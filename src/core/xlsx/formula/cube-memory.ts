// In-memory reference CubeProvider for local OLAP/pivot cubes and testing.
import type { CellError } from '../model';
import type {
	CubeKpiProperty,
	CubeMemberHandle,
	CubeMemberResult,
	CubeProvider,
	CubeSetHandle,
	CubeSetOptions,
	CubeSetResult,
	CubeValueArg,
} from './providers';
import { ERR, type Scalar } from './values';

export interface MemoryCubeMember {
	readonly caption?: string;
	readonly properties?: Readonly<Record<string, Scalar>>;
}

export interface MemoryCubeKpiEntry {
	readonly expression: string;
	readonly caption?: string;
	readonly value?: Scalar;
}

export interface MemoryCubeSetEntry {
	readonly expression: string;
	readonly caption?: string;
	readonly sortValue?: number | string;
}

export interface MemoryCubeData {
	readonly members?: Readonly<Record<string, MemoryCubeMember>>;
	readonly kpis?: Readonly<
		Record<string, Readonly<Partial<Record<CubeKpiProperty, MemoryCubeKpiEntry>>>>
	>;
	readonly sets?: Readonly<Record<string, readonly MemoryCubeSetEntry[]>>;
	/**
	 * Fact values keyed by `expressions.join('|')` (exact order or sorted order), or a resolver
	 * function receiving the flattened list of member expressions.
	 */
	readonly values?: Readonly<Record<string, Scalar>> | ((expressions: readonly string[]) => Scalar);
}

const defaultCaption = (expression: string): string => {
	const match = /\[([^\]]+)\]\s*$/u.exec(expression);
	return match?.[1] ?? expression;
};

const compareSort = (a: number | string | undefined, b: number | string | undefined): number => {
	if (typeof a === 'number' && typeof b === 'number') return a - b;
	return String(a ?? '').localeCompare(String(b ?? ''));
};

/** Creates an in-memory `CubeProvider` backed by named cubes. */
export function createMemoryCubeProvider(
	cubes: Readonly<Record<string, MemoryCubeData>>,
): CubeProvider {
	const resolvedSets = new Map<string, readonly MemoryCubeSetEntry[]>();
	let nextSetId = 1;

	const getCube = (connection: string): MemoryCubeData | undefined => cubes[connection];

	const resolveEntries = (
		cube: MemoryCubeData,
		setExpression: string,
		options?: CubeSetOptions,
	): readonly MemoryCubeSetEntry[] | undefined => {
		const raw = cube.sets?.[setExpression];
		if (!raw) return undefined;
		const entries = [...raw];
		const order = options?.sortOrder ?? 0;
		if (order === 1 || order === 2) {
			const dir = order === 1 ? 1 : -1;
			entries.sort((a, b) => dir * compareSort(a.sortValue, b.sortValue));
		} else if (order === 3 || order === 4) {
			const dir = order === 3 ? 1 : -1;
			entries.sort(
				(a, b) =>
					dir *
					(a.caption ?? defaultCaption(a.expression)).localeCompare(
						b.caption ?? defaultCaption(b.expression),
					),
			);
		} else if (order === 6) {
			entries.reverse();
		}
		return entries;
	};

	const lookupValue = (cube: MemoryCubeData, expressions: readonly string[]): Scalar => {
		if (typeof cube.values === 'function') return cube.values(expressions);
		if (!cube.values) return ERR.NA;
		const direct = cube.values[expressions.join('|')];
		if (direct !== undefined) return direct;
		const sorted = cube.values[[...expressions].sort().join('|')];
		if (sorted !== undefined) return sorted;
		return ERR.NA;
	};

	return {
		member(
			connection: string,
			memberExpression: string | readonly string[],
			caption?: string,
		): CubeMemberResult | CellError {
			const cube = getCube(connection);
			if (!cube) return ERR.NA;
			const key = Array.isArray(memberExpression)
				? memberExpression.join('|')
				: (memberExpression as string);
			const entry = cube.members?.[key];
			if (!entry) return ERR.NA;
			return {
				token: key,
				caption: caption ?? entry.caption ?? defaultCaption(key),
			};
		},

		kpiMember(
			connection: string,
			kpiName: string,
			kpiProperty: CubeKpiProperty,
			caption?: string,
		): CubeMemberResult | CellError {
			const cube = getCube(connection);
			const entry = cube?.kpis?.[kpiName]?.[kpiProperty];
			if (!entry) return ERR.NA;
			return {
				token: entry.expression,
				caption: caption ?? entry.caption ?? defaultCaption(entry.expression),
			};
		},

		memberProperty(
			connection: string,
			member: string | CubeMemberHandle,
			property: string,
		): Scalar {
			const cube = getCube(connection);
			if (!cube) return ERR.NA;
			const key = typeof member === 'string' ? member : member.token;
			const value = cube.members?.[key]?.properties?.[property];
			return value !== undefined ? value : ERR.NA;
		},

		set(
			connection: string,
			setExpression: string,
			options?: CubeSetOptions,
		): CubeSetResult | CellError {
			const cube = getCube(connection);
			if (!cube) return ERR.NA;
			const entries = resolveEntries(cube, setExpression, options);
			if (!entries) return ERR.NA;
			const token = `set:${nextSetId++}:${connection}:${setExpression}`;
			resolvedSets.set(token, entries);
			return { token, caption: options?.caption ?? '' };
		},

		setCount(set: CubeSetHandle): number | CellError {
			const entries = resolvedSets.get(set.token);
			return entries ? entries.length : ERR.NA;
		},

		rankedMember(
			connection: string,
			set: string | CubeSetHandle,
			rank: number,
			caption?: string,
		): CubeMemberResult | CellError {
			const cube = getCube(connection);
			if (!cube) return ERR.NA;
			const entries = typeof set === 'string' ? cube.sets?.[set] : resolvedSets.get(set.token);
			if (!entries) return ERR.NA;
			const item = entries[rank - 1];
			if (!item) return ERR.NA;
			return {
				token: item.expression,
				caption: caption ?? item.caption ?? defaultCaption(item.expression),
			};
		},

		value(connection: string, expressions: readonly CubeValueArg[]): Scalar {
			const cube = getCube(connection);
			if (!cube) return ERR.NA;
			const flat: string[] = [];
			for (const expr of expressions) {
				if (typeof expr === 'string') {
					flat.push(expr);
				} else if (expr.kind === 'member') {
					flat.push(expr.token);
				} else {
					const entries = resolvedSets.get(expr.token);
					if (!entries) return ERR.NA;
					for (const item of entries) flat.push(item.expression);
				}
			}
			// Also check KPI direct values when a single KPI member token is requested.
			if (flat.length === 1 && cube.kpis) {
				const token = flat[0] as string;
				for (const kpi of Object.values(cube.kpis)) {
					for (const entry of Object.values(kpi)) {
						if (entry && entry.expression === token && entry.value !== undefined) {
							return entry.value;
						}
					}
				}
			}
			return lookupValue(cube, flat);
		},
	};
}
