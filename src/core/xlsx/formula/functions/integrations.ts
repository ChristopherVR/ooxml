// Host-backed external/integration functions: COPILOT, STOCKHISTORY, TRANSLATE,
// DETECTLANGUAGE, IMAGE, RTD and WEBSERVICE.
import type { CallContext } from '../context';
import type {
	CalcExternalProviders,
	CopilotPromptPart,
	ImageRequest,
	ImageSizing,
	StockHistoryBar,
	StockHistoryInterval,
	StockHistoryProperty,
} from '../providers';
import { ERR, fail, isError, Matrix, RefValue, type Scalar, type Value } from '../values';
import { int, num, omitted, spec, str } from './helpers';
import type { FunctionSpec } from './types';

const MAX_WEBSERVICE_URL = 2048;
const MAX_CELL_TEXT = 32_767;
const STOCK_HEADERS = ['Date', 'Close', 'Open', 'High', 'Low', 'Volume'] as const;

function integrationSpec(
	provider: keyof CalcExternalProviders,
	name: string,
	category: string,
	syntax: string,
	description: string,
	minArgs: number,
	maxArgs: number,
	fn: NonNullable<FunctionSpec['fn']>,
	options?: { params?: FunctionSpec['params']; volatile?: boolean },
): FunctionSpec {
	return {
		...spec(name, category, syntax, description, minArgs, maxArgs, fn, options?.params),
		provider,
		...(options?.volatile ? { volatile: true } : {}),
	};
}

function parseHttpUrl(raw: string, onDisallowedScheme: () => never): string {
	const trimmed = raw.trim();
	if (trimmed.length === 0) fail(ERR.VALUE);
	let parsed: URL;
	try {
		parsed = new URL(trimmed);
	} catch {
		fail(ERR.VALUE);
	}
	if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') onDisallowedScheme();
	return trimmed;
}

function toCopilotContext(ctx: CallContext, arg: Value): Scalar | Scalar[][] {
	if (arg instanceof RefValue || arg instanceof Matrix) {
		const m = ctx.toMatrix(arg).block();
		return m.data.map((row) => [...row]);
	}
	return ctx.toScalar(arg);
}

function stockPropertyValue(bar: StockHistoryBar, prop: StockHistoryProperty): Scalar {
	switch (prop) {
		case 0:
			return bar.date;
		case 1:
			return bar.close ?? null;
		case 2:
			return bar.open ?? null;
		case 3:
			return bar.high ?? null;
		case 4:
			return bar.low ?? null;
		case 5:
			return bar.volume ?? null;
	}
}

function evalStockHistory(args: Value[], ctx: CallContext): Value {
	const provider = ctx.frame.host.providers?.stockHistory;
	if (!provider) fail(ERR.NAME);
	const stock = str(args[0]).trim();
	if (stock.length === 0) fail(ERR.VALUE);
	const startDate = Math.trunc(num(args[1]));
	if (startDate < 0) fail(ERR.VALUE);
	const endDate = omitted(args, 2) ? startDate : Math.trunc(num(args[2]));
	if (endDate < 0 || endDate < startDate) fail(ERR.VALUE);
	const intervalRaw = omitted(args, 3) ? 0 : int(args[3]);
	if (intervalRaw < 0 || intervalRaw > 2) fail(ERR.VALUE);
	const headers = omitted(args, 4) ? 1 : int(args[4]);
	if (headers < 0 || headers > 2) fail(ERR.VALUE);
	const properties: StockHistoryProperty[] = [];
	for (let i = 5; i < args.length; i++) {
		if (omitted(args, i)) continue;
		const p = int(args[i]);
		if (p < 0 || p > 5) fail(ERR.VALUE);
		properties.push(p as StockHistoryProperty);
	}
	if (properties.length === 0) properties.push(0, 1);
	const result = provider.history({
		stock,
		startDate,
		endDate,
		interval: intervalRaw as StockHistoryInterval,
		date1904: ctx.date1904,
	});
	if (isError(result)) return result;
	const rows: Scalar[][] = [];
	if (headers === 2) {
		rows.push([stock, ...Array<Scalar>(properties.length - 1).fill('')]);
	}
	if (headers === 1 || headers === 2) {
		rows.push(properties.map((p) => STOCK_HEADERS[p]));
	}
	for (const bar of result) {
		rows.push(properties.map((p) => stockPropertyValue(bar, p)));
	}
	if (rows.length === 0) return ERR.NA;
	if (rows.length === 1 && rows[0]?.length === 1) return rows[0][0] ?? null;
	return new Matrix(rows);
}

function evalImage(args: Value[], ctx: CallContext): Value {
	const provider = ctx.frame.host.providers?.image;
	if (!provider) fail(ERR.NAME);
	const source = parseHttpUrl(str(args[0]), () => fail(ERR.BLOCKED));
	const altText = omitted(args, 1) ? '' : str(args[1]);
	const sizingRaw = omitted(args, 2) ? 0 : int(args[2]);
	if (sizingRaw < 0 || sizingRaw > 3) fail(ERR.VALUE);
	const sizing = sizingRaw as ImageSizing;
	const hasHeight = !omitted(args, 3);
	const hasWidth = !omitted(args, 4);
	let height: number | undefined;
	let width: number | undefined;
	if (sizing === 3) {
		if (!hasHeight || !hasWidth) fail(ERR.VALUE);
		height = num(args[3]);
		width = num(args[4]);
		if (!Number.isFinite(height) || !Number.isFinite(width) || height <= 0 || width <= 0)
			fail(ERR.VALUE);
	} else if (hasHeight || hasWidth) {
		fail(ERR.VALUE);
	}
	const req: ImageRequest = {
		source,
		altText,
		sizing,
		...(height !== undefined ? { height } : {}),
		...(width !== undefined ? { width } : {}),
	};
	return provider.resolve(req);
}

export const INTEGRATION_FUNCTIONS: FunctionSpec[] = [
	integrationSpec(
		'copilot',
		'COPILOT',
		'Lookup & Reference',
		'COPILOT(prompt_part1, [context1], [prompt_part2], [context2], ...)',
		'Calls a host-supplied AI provider with prompt text and optional cell or range context.',
		1,
		255,
		(args, ctx) => {
			const provider = ctx.frame.host.providers?.copilot;
			if (!provider) fail(ERR.NAME);
			const parts: CopilotPromptPart[] = [];
			for (let i = 0; i < args.length; i += 2) {
				const promptScalar = ctx.toScalar(args[i] ?? null);
				if (isError(promptScalar)) return promptScalar;
				const prompt = str(promptScalar);
				if (i === 0 && prompt.trim().length === 0) fail(ERR.VALUE);
				const contextArg = args[i + 1];
				const part: CopilotPromptPart =
					contextArg !== undefined && contextArg !== null
						? { prompt, context: toCopilotContext(ctx, contextArg) }
						: { prompt };
				parts.push(part);
			}
			const out = provider.evaluate({ parts });
			if (Array.isArray(out)) {
				if (out.length === 0 || (out[0]?.length ?? 0) === 0) return ERR.CALC;
				return new Matrix(out.map((row) => [...row]));
			}
			return out as Scalar;
		},
		{ params: ['any'] },
	),
	integrationSpec(
		'stockHistory',
		'STOCKHISTORY',
		'Financial',
		'STOCKHISTORY(stock, start_date, [end_date], [interval], [headers], [property0], ...)',
		'Retrieves historical market data from a host-supplied provider.',
		2,
		11,
		evalStockHistory,
	),
	integrationSpec(
		'translate',
		'TRANSLATE',
		'Text',
		'TRANSLATE(text, [source_language], [target_language])',
		'Translates text using a host-supplied translation provider.',
		1,
		3,
		(args, ctx) => {
			const provider = ctx.frame.host.providers?.translate;
			if (!provider) fail(ERR.NAME);
			const text = str(args[0]);
			const source = omitted(args, 1) ? undefined : str(args[1]).trim();
			const target = omitted(args, 2) ? undefined : str(args[2]).trim();
			if (source !== undefined && source.length === 0) fail(ERR.VALUE);
			if (target !== undefined && target.length === 0) fail(ERR.VALUE);
			return provider.translate(text, source, target);
		},
	),
	integrationSpec(
		'detectLanguage',
		'DETECTLANGUAGE',
		'Text',
		'DETECTLANGUAGE(text)',
		'Identifies the language of text using a host-supplied provider.',
		1,
		1,
		(args, ctx) => {
			const provider = ctx.frame.host.providers?.detectLanguage;
			if (!provider) fail(ERR.NAME);
			return provider.detectLanguage(str(args[0]));
		},
	),
	integrationSpec(
		'image',
		'IMAGE',
		'Lookup & Reference',
		'IMAGE(source, [alt_text], [sizing], [height], [width])',
		'Resolves an image URL through a host-supplied provider.',
		1,
		5,
		evalImage,
	),
	integrationSpec(
		'rtd',
		'RTD',
		'Lookup & Reference',
		'RTD(progID, server, topic1, [topic2], ...)',
		'Retrieves real-time data from a host-supplied provider.',
		3,
		255,
		(args, ctx) => {
			const provider = ctx.frame.host.providers?.rtd;
			if (!provider) fail(ERR.NAME);
			const progId = str(args[0]).trim();
			if (progId.length === 0) fail(ERR.VALUE);
			const server = omitted(args, 1) ? '' : str(args[1]);
			const topics: string[] = [];
			for (let i = 2; i < args.length; i++) topics.push(str(args[i]));
			return provider.read(progId, server, topics);
		},
		{ volatile: true },
	),
	integrationSpec(
		'webservice',
		'WEBSERVICE',
		'Web',
		'WEBSERVICE(url)',
		'Retrieves data from an HTTP(S) URL through a host-supplied provider.',
		1,
		1,
		(args, ctx) => {
			const provider = ctx.frame.host.providers?.webservice;
			if (!provider) fail(ERR.NAME);
			const url = parseHttpUrl(str(args[0]), () => fail(ERR.VALUE));
			if (url.length > MAX_WEBSERVICE_URL) fail(ERR.VALUE);
			const res = provider.fetch(url);
			if (isError(res)) return res;
			if (res.length > MAX_CELL_TEXT) return ERR.VALUE;
			return res;
		},
	),
];
