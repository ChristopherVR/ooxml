// Host-backed OLAP cube functions: CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY,
// CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT and CUBEVALUE.
import type { CallContext, EvalHost } from '../context';
import type {
	CubeHandle,
	CubeKpiProperty,
	CubeMemberHandle,
	CubeProvider,
	CubeSetHandle,
	CubeSetOptions,
	CubeSortOrder,
	CubeValueArg,
} from '../providers';
import { ERR, fail, isError, Matrix, RefValue, type Value } from '../values';
import { int, omitted, scalar, spec, str } from './helpers';
import type { FunctionSpec } from './types';

const C = 'Cube';

interface CubeHandleStore {
	readonly byCell: Map<string, CubeHandle>;
	readonly byCaption: Map<string, CubeHandle>;
}

const HANDLES = new WeakMap<EvalHost, CubeHandleStore>();

function handleStore(host: EvalHost): CubeHandleStore {
	let store = HANDLES.get(host);
	if (!store) {
		store = { byCell: new Map(), byCaption: new Map() };
		HANDLES.set(host, store);
	}
	return store;
}

const cellTag = (sheet: number, row: number, col: number): string => `${sheet}:${row}:${col}`;

function recordHandle(ctx: CallContext, handle: CubeHandle): string {
	const store = handleStore(ctx.frame.host);
	store.byCell.set(cellTag(ctx.sheet, ctx.row, ctx.col), handle);
	store.byCaption.set(`${handle.kind}:${handle.caption}`, handle);
	return handle.caption;
}

function resolveHandle(
	ctx: CallContext,
	arg: Value | undefined,
	kind?: 'member' | 'set',
): CubeHandle | undefined {
	if (arg === undefined) return undefined;
	const store = handleStore(ctx.frame.host);
	if (arg instanceof RefValue && arg.isCell()) {
		const area = arg.first;
		if (area) {
			// Ensure the referenced cell's formula has evaluated first.
			const cellVal = ctx.readCell(area.sheet, area.range.start.row, area.range.start.col);
			if (isError(cellVal)) fail(cellVal);
			const found = store.byCell.get(
				cellTag(area.sheet, area.range.start.row, area.range.start.col),
			);
			if (found && (!kind || found.kind === kind)) return found;
		}
	}
	const s = ctx.toScalar(arg);
	if (isError(s)) fail(s);
	if (typeof s === 'string') {
		if (kind) return store.byCaption.get(`${kind}:${s}`);
		return store.byCaption.get(`member:${s}`) ?? store.byCaption.get(`set:${s}`);
	}
	return undefined;
}

function cubeProvider(ctx: CallContext): CubeProvider {
	const provider = ctx.frame.host.providers?.cube;
	if (!provider) fail(ERR.NAME);
	return provider;
}

function requireText(arg: Value | undefined): string {
	const text = str(arg);
	if (text.length === 0) fail(ERR.VALUE);
	return text;
}

function optCaption(args: Value[], index: number): string | undefined {
	return omitted(args, index) ? undefined : str(args[index]);
}

function cubeSpec(
	name: string,
	syntax: string,
	description: string,
	minArgs: number,
	maxArgs: number,
	fn: NonNullable<FunctionSpec['fn']>,
	params?: FunctionSpec['params'],
): FunctionSpec {
	return { ...spec(name, C, syntax, description, minArgs, maxArgs, fn, params), provider: 'cube' };
}

function extractExpressions(ctx: CallContext, arg: Value): string | string[] {
	if (arg instanceof RefValue || arg instanceof Matrix) {
		const values = ctx.toMatrix(arg).flat();
		const out: string[] = [];
		for (const v of values) {
			if (isError(v)) fail(v);
			if (v !== null) out.push(str(v));
		}
		if (out.length === 0) fail(ERR.VALUE);
		return out.length === 1 ? (out[0] as string) : out;
	}
	return requireText(arg);
}

function collectCubeValueArgs(ctx: CallContext, args: Value[]): CubeValueArg[] {
	const out: CubeValueArg[] = [];
	for (let i = 1; i < args.length; i++) {
		const arg = args[i];
		if (arg === undefined || arg === null) continue;
		if (arg instanceof RefValue && !arg.isCell()) {
			const area = arg.first;
			if (!area || arg.areas.length > 1) fail(ERR.VALUE);
			for (let r = area.range.start.row; r <= area.range.end.row; r++) {
				for (let c = area.range.start.col; c <= area.range.end.col; c++) {
					const single = new RefValue([
						{ sheet: area.sheet, range: { start: { row: r, col: c }, end: { row: r, col: c } } },
					]);
					const handle = resolveHandle(ctx, single);
					if (handle) {
						out.push(handle);
						continue;
					}
					const v = ctx.readCell(area.sheet, r, c);
					if (isError(v)) fail(v);
					if (v !== null && v !== '') out.push(str(v));
				}
			}
			continue;
		}
		if (arg instanceof Matrix) {
			for (const v of arg.flat()) {
				if (isError(v)) fail(v);
				if (v !== null && v !== '') out.push(str(v));
			}
			continue;
		}
		const handle = resolveHandle(ctx, arg);
		if (handle) {
			out.push(handle);
			continue;
		}
		const s = ctx.toScalar(arg);
		if (isError(s)) fail(s);
		if (s !== null && s !== '') out.push(str(s));
	}
	return out;
}

export const CUBE_FUNCTIONS: FunctionSpec[] = [
	cubeSpec(
		'CUBEKPIMEMBER',
		'CUBEKPIMEMBER(connection, kpi_name, kpi_property, [caption])',
		'Returns a key performance indicator (KPI) property from a cube.',
		3,
		4,
		(args, ctx) => {
			const connection = requireText(args[0]);
			const kpiName = requireText(args[1]);
			const prop = int(args[2]);
			if (prop < 1 || prop > 6) fail(ERR.VALUE);
			const res = cubeProvider(ctx).kpiMember(
				connection,
				kpiName,
				prop as CubeKpiProperty,
				optCaption(args, 3),
			);
			if (isError(res)) return res;
			return recordHandle(ctx, {
				kind: 'member',
				connection,
				token: res.token,
				caption: res.caption,
			});
		},
	),
	cubeSpec(
		'CUBEMEMBER',
		'CUBEMEMBER(connection, member_expression, [caption])',
		'Returns a member or tuple from a cube.',
		2,
		3,
		(args, ctx) => {
			const connection = requireText(args[0]);
			const expr = extractExpressions(ctx, args[1] ?? null);
			const res = cubeProvider(ctx).member(connection, expr, optCaption(args, 2));
			if (isError(res)) return res;
			return recordHandle(ctx, {
				kind: 'member',
				connection,
				token: res.token,
				caption: res.caption,
			});
		},
		['value', 'any', 'value'],
	),
	cubeSpec(
		'CUBEMEMBERPROPERTY',
		'CUBEMEMBERPROPERTY(connection, member_expression, property)',
		'Returns the value of a member property from a cube.',
		3,
		3,
		(args, ctx) => {
			const connection = requireText(args[0]);
			const property = requireText(args[2]);
			const handle = resolveHandle(ctx, args[1], 'member') as CubeMemberHandle | undefined;
			const member: string | CubeMemberHandle =
				handle ?? requireText(ctx.toScalar(args[1] ?? null));
			return cubeProvider(ctx).memberProperty(connection, member, property);
		},
		['value', 'any', 'value'],
	),
	cubeSpec(
		'CUBERANKEDMEMBER',
		'CUBERANKEDMEMBER(connection, set_expression, rank, [caption])',
		'Returns the nth, or ranked, member in a set.',
		3,
		4,
		(args, ctx) => {
			const connection = requireText(args[0]);
			const rank = int(args[2]);
			if (rank < 1) fail(ERR.VALUE);
			const handle = resolveHandle(ctx, args[1], 'set') as CubeSetHandle | undefined;
			const set: string | CubeSetHandle = handle ?? requireText(ctx.toScalar(args[1] ?? null));
			const res = cubeProvider(ctx).rankedMember(connection, set, rank, optCaption(args, 3));
			if (isError(res)) return res;
			return recordHandle(ctx, {
				kind: 'member',
				connection,
				token: res.token,
				caption: res.caption,
			});
		},
		['value', 'any', 'value', 'value'],
	),
	cubeSpec(
		'CUBESET',
		'CUBESET(connection, set_expression, [caption], [sort_order], [sort_by])',
		'Defines a calculated set of members or tuples from a cube.',
		2,
		5,
		(args, ctx) => {
			const connection = requireText(args[0]);
			const setExpression = requireText(args[1]);
			const caption = optCaption(args, 2);
			let sortOrder: CubeSortOrder | undefined;
			if (!omitted(args, 3)) {
				const n = int(args[3]);
				if (n < 0 || n > 6) fail(ERR.VALUE);
				sortOrder = n as CubeSortOrder;
			}
			const sortBy = omitted(args, 4) ? undefined : str(args[4]);
			if ((sortOrder === 1 || sortOrder === 2) && (!sortBy || sortBy.length === 0)) fail(ERR.VALUE);
			const options: CubeSetOptions = {
				...(caption !== undefined ? { caption } : {}),
				...(sortOrder !== undefined ? { sortOrder } : {}),
				...(sortBy !== undefined ? { sortBy } : {}),
			};
			const res = cubeProvider(ctx).set(connection, setExpression, options);
			if (isError(res)) return res;
			return recordHandle(ctx, {
				kind: 'set',
				connection,
				token: res.token,
				caption: res.caption,
			});
		},
	),
	cubeSpec(
		'CUBESETCOUNT',
		'CUBESETCOUNT(set)',
		'Returns the number of items in a set.',
		1,
		1,
		(args, ctx) => {
			const handle = resolveHandle(ctx, args[0], 'set') as CubeSetHandle | undefined;
			if (!handle) {
				const s = scalar(ctx.toScalar(args[0] ?? null));
				if (isError(s)) return s;
				fail(ERR.VALUE);
			}
			return cubeProvider(ctx).setCount(handle);
		},
		['any'],
	),
	cubeSpec(
		'CUBEVALUE',
		'CUBEVALUE(connection, [member_expression1], ...)',
		'Returns an aggregated value from a cube.',
		1,
		255,
		(args, ctx) => {
			const connection = requireText(args[0]);
			const expressions = collectCubeValueArgs(ctx, args);
			return cubeProvider(ctx).value(connection, expressions);
		},
		['value', 'any'],
	),
];
