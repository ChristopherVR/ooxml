/**
 * A bounded repair of chart parts this library once saved malformed. Before the c15/c16 writer fix,
 * hiding a series or editing a data-label range wrote the namespace declarations as elements
 * (`<xmlns:c15>uri</xmlns:c15>`), used the `c15`/`c16` prefixes undeclared, and wrote a hidden
 * series' literal name as `c15:v`. The strict DOM parser (`chart/parseChartSpace`) rejects such a
 * part; the pptx load path calls {@link repairChartPartPrefixes} once and retries. The neutral
 * parser itself stays strict, and the save path writes well-formed parts.
 *
 * @module chart-part-repair
 */
import { parseChartSpace, type ChartParseIssue, type ChartSpace } from '../../../chart/index';

/** The standard namespaces of the Office chart prefixes the repair may declare. */
export const OFFICE_CHART_PREFIXES: Readonly<Record<string, string>> = {
	c14: 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
	c15: 'http://schemas.microsoft.com/office/drawing/2012/chart',
	c16: 'http://schemas.microsoft.com/office/drawing/2014/chart',
	c16r2: 'http://schemas.microsoft.com/office/drawing/2015/06/chart',
	c16r3: 'http://schemas.microsoft.com/office/drawing/2017/03/chart',
	cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
	a16: 'http://schemas.microsoft.com/office/drawing/2014/main',
	mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
	xr: 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
};

/** What {@link repairChartPartPrefixes} changed. */
export interface ChartPartRepair {
	text: string;
	/** The prefixes declared on the root. */
	declared: string[];
	/** Whether `xmlns:*` elements were removed or `c15:v` was mapped back to `c:v`. */
	rewroteOwnOutput: boolean;
}

const ROOT_START = /<([A-Za-z_][\w.-]*:)?chartSpace\b[^>]*>/u;
const XMLNS_ELEMENT = /<xmlns:[\w.-]+\s*(?:\/>|>[^<]*<\/xmlns:[\w.-]+\s*>)/gu;
const C15_VALUE = /<(\/?)c15:v(\s*\/?)>/gu;

/**
 * The part with the library's old malformed output rewritten (`xmlns:*` elements removed, `c15:v`
 * as `c:v`) and every known Office prefix it uses but the root does not declare declared on the
 * root; undefined when there is nothing to repair. Unknown prefixes are left alone, so a part that
 * is malformed some other way stays malformed.
 */
export function repairChartPartPrefixes(text: string): ChartPartRepair | undefined {
	let body = text.replace(XMLNS_ELEMENT, '');
	body = body.replace(C15_VALUE, '<$1c:v$2>');
	const rewroteOwnOutput = body !== text;
	const root = ROOT_START.exec(body);
	if (!root) return rewroteOwnOutput ? { text: body, declared: [], rewroteOwnOutput } : undefined;
	const declaredOnRoot = new Set(
		[...root[0].matchAll(/\sxmlns:([\w.-]+)\s*=/gu)].map((match) => match[1]),
	);
	const used = new Set(
		[...body.matchAll(/<\/?([A-Za-z_][\w.-]*):[\w.-]+|\s([A-Za-z_][\w.-]*):[\w.-]+\s*=/gu)]
			.map((match) => match[1] ?? match[2])
			.filter((prefix): prefix is string => prefix !== undefined && prefix !== 'xmlns'),
	);
	const declared = [...used].filter(
		(prefix) => !declaredOnRoot.has(prefix) && Object.hasOwn(OFFICE_CHART_PREFIXES, prefix),
	);
	if (declared.length === 0 && !rewroteOwnOutput) return undefined;
	const declarations = declared
		.map((prefix) => ` xmlns:${prefix}="${OFFICE_CHART_PREFIXES[prefix]}"`)
		.join('');
	const end = root.index + root[0].length - (root[0].endsWith('/>') ? 2 : 1);
	return {
		text: body.slice(0, end) + declarations + body.slice(end),
		declared,
		rewroteOwnOutput,
	};
}

/** What {@link readChartPartModel} returns. */
export interface ChartPartModel {
	/** The neutral model; undefined for ChartEx and unreadable parts. */
	chartSpace?: ChartSpace;
	/** `CHART_PREFIX_REPAIRED` when the part only parsed after the repair. */
	repairIssue?: ChartParseIssue;
}

function parseClassic(text: string): ChartSpace | undefined {
	const { chartSpace, issues } = parseChartSpace(text);
	return issues.some((issue) => issue.code === 'CHART_ROOT_UNEXPECTED') ? undefined : chartSpace;
}

/**
 * The neutral model of a chart part as the pptx loader reads it: parsed strictly, and when that
 * fails, repaired once ({@link repairChartPartPrefixes}) and parsed again. A part still malformed
 * after the repair is unreadable.
 */
export function readChartPartModel(text: string): ChartPartModel {
	try {
		return { chartSpace: parseClassic(text) };
	} catch {
		const repair = repairChartPartPrefixes(text);
		if (!repair) return {};
		try {
			const chartSpace = parseClassic(repair.text);
			const what = [
				...(repair.declared.length ? [`declared ${repair.declared.join(', ')}`] : []),
				...(repair.rewroteOwnOutput ? ['rewrote xmlns elements and c15:v'] : []),
			].join('; ');
			return {
				...(chartSpace ? { chartSpace } : {}),
				repairIssue: {
					code: 'CHART_PREFIX_REPAIRED',
					message: `The chart part was not namespace well-formed and was repaired to load (${what}).`,
				},
			};
		} catch {
			return {};
		}
	}
}
