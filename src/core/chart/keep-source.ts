// When a chart part may be rewritten from its edited model. `writeChartSpace` plus
// `withUnmodelledChartChildren` is lossless only when everything the parser reported is one of the
// top-level children that are carried over; a part with anything else unmodelled (inside the plot
// area, a series or an axis) or with a value the parser could not read would lose it. Such a part
// keeps its source bytes: only reference moves (rows inserted, a sheet renamed) are patched into it,
// formula by formula, the way the DOM patch path always did; any other edit must be refused by the
// caller.
import { NS, buildXml, elements, first, parseXml, type XmlElement } from '../xml/index';
import type { ChartSpace } from './model';
import { parseChartSpace } from './parse-space';
import { writeChartSpace, type WriteChartSpaceOptions } from './write-space';
import { withUnmodelledChartChildren } from './keep-unmodelled';
import type { ChartDataSource, ChartParseIssue } from './model-series';

/** The unmodelled elements `withUnmodelledChartChildren` carries into an edited part. */
const CARRIED = new Set([
	'c:chartSpace/c:clrMapOvr',
	'c:chartSpace/c:pivotSource',
	'c:chartSpace/c:protection',
	'c:chart/c:pivotFmts',
]);

/** The issues an edited part would lose; empty when the part can be written from its model. */
export function chartIssuesLostOnWrite(issues: readonly ChartParseIssue[]): ChartParseIssue[] {
	return issues.filter(
		(issue) =>
			issue.code !== 'CHART_ELEMENT_NOT_MODELLED' ||
			issue.element === undefined ||
			!CARRIED.has(issue.element),
	);
}

/** Series data sources in a fixed order: name, categories, values, x, y, bubble sizes. */
const SOURCES = ['tx', 'categories', 'values', 'xValues', 'yValues', 'bubbleSizes'] as const;
const ELEMENT: Record<(typeof SOURCES)[number], string> = {
	tx: 'tx',
	categories: 'cat',
	values: 'val',
	xValues: 'xVal',
	yValues: 'yVal',
	bubbleSizes: 'bubbleSize',
};

const sourceOf = (
	series: ChartSpace['plotArea']['groups'][number]['series'][number],
	key: (typeof SOURCES)[number],
): ChartDataSource | undefined => (key === 'tx' ? series.tx?.reference : series[key]);

/** `space` with every series formula replaced by `formula(groupIndex, seriesIndex, key)`. */
function withFormulas(
	space: ChartSpace,
	formula: (g: number, s: number, key: (typeof SOURCES)[number]) => string | undefined,
): ChartSpace {
	const copy = structuredClone(space);
	copy.plotArea.groups.forEach((group, g) =>
		group.series.forEach((series, s) => {
			for (const key of SOURCES) {
				const source = sourceOf(series, key);
				const next = formula(g, s, key);
				if (source && next !== undefined) source.formula = next;
			}
		}),
	);
	return copy;
}

const stable = (value: unknown): unknown =>
	Array.isArray(value)
		? value.map(stable)
		: value && typeof value === 'object'
			? Object.fromEntries(
					Object.keys(value)
						.sort()
						.map((key) => [key, stable((value as Record<string, unknown>)[key])]),
				)
			: value;
const same = (a: unknown, b: unknown) => JSON.stringify(stable(a)) === JSON.stringify(stable(b));

/**
 * `source` with the series formulas of `current` written into its `c:f` elements, when `current`
 * differs from `loaded` (the model of `source`) in those formulas only; `undefined` otherwise (or
 * when nothing changed).
 */
export function patchChartFormulas(
	source: string,
	loaded: ChartSpace,
	current: ChartSpace,
): string | undefined {
	const groupsOf = (space: ChartSpace) => space.plotArea.groups;
	if (groupsOf(loaded).length !== groupsOf(current).length) return undefined;
	const formulaAt =
		(space: ChartSpace) => (g: number, s: number, key: (typeof SOURCES)[number]) => {
			const series = groupsOf(space)[g]?.series[s];
			return series ? sourceOf(series, key)?.formula : undefined;
		};
	// Only the formulas may differ: give `current` the loaded formulas and compare.
	if (!same(withFormulas(current, formulaAt(loaded)), loaded)) return undefined;
	const doc = parseXml(source, { label: 'Chart part' });
	const plotArea = first(first(doc.documentElement, 'chart', NS.c), 'plotArea', NS.c);
	const groupElements = plotArea
		? elements(plotArea).filter(
				(node) => node.namespaceURI === NS.c && /Chart$/.test(node.localName),
			)
		: [];
	let changed = false;
	groupElements.forEach((groupElement, g) =>
		elementsNamed(groupElement, 'ser').forEach((ser, s) => {
			for (const key of SOURCES) {
				const next = formulaAt(current)(g, s, key);
				const holder = first(ser, ELEMENT[key], NS.c);
				const reference = holder && elements(holder).find((node) => /Ref$/.test(node.localName));
				const f = first(reference, 'f', NS.c);
				if (f && next !== undefined && f.textContent !== next) {
					f.textContent = next;
					changed = true;
				}
			}
		}),
	);
	if (!changed) return undefined;
	const declaration = /^<\?xml[^>]*\?>(\r\n|\n)?/.exec(source)?.[0] ?? '';
	return declaration + buildXml(doc.documentElement);
}

const elementsNamed = (parent: XmlElement, local: string) =>
	elements(parent).filter((node) => node.namespaceURI === NS.c && node.localName === local);

/**
 * Writes an edited source-backed chart only when the parser can preserve every reported
 * element. Unchanged models return the original XML, including its lexical formatting.
 * Carried top-level children are restored at their schema positions.
 */
export function writeChartSpaceOverSource(
	source: string,
	space: ChartSpace,
	options?: WriteChartSpaceOptions,
): string {
	const parsed = parseChartSpace(source);
	if (same(parsed.chartSpace, space)) return source;
	if (chartIssuesLostOnWrite(parsed.issues).length)
		throw new Error('This chart has content that cannot be preserved through an edit.');
	return withUnmodelledChartChildren(
		source,
		writeChartSpace(space, options),
		space.protection === null,
	);
}
