// The axes of a chart group after a type switch: the axes Excel writes for a new chart (lifted from
// xlsx `write/chart-space.ts`, see PROVENANCE.md), and the conversion of a category axis into the
// value axis a scatter or bubble chart plots its x values on, and back.
import type { ChartAxis, ChartAxisKind } from '../model';

/** The axis ids Excel gives the two axes of a new chart. */
export const DEFAULT_CHART_AXIS_IDS = [500000001, 500000002] as const;

/** The category (or X value) axis and the value axis of a new chart, crossing at zero. */
export function defaultChartAxes(
	horizontal: boolean,
	scatter: boolean,
	ids: readonly [number, number] = DEFAULT_CHART_AXIS_IDS,
): [ChartAxis, ChartAxis] {
	const [firstId, secondId] = ids;
	// A function, so the two axes never share a scaling or format object.
	const common = () => ({
		scaling: { orientation: 'minMax' },
		deleted: false,
		minorGridlines: false,
		numberFormat: { formatCode: 'General', sourceLinked: true },
		majorTickMark: 'out',
		minorTickMark: 'none',
		tickLabelPosition: 'nextTo',
		crosses: 'autoZero',
	});
	const first: ChartAxis = scatter
		? {
				...common(),
				kind: 'val',
				id: firstId,
				position: 'b',
				majorGridlines: false,
				crossAxisId: secondId,
				crossBetween: 'midCat',
			}
		: {
				...common(),
				kind: 'cat',
				id: firstId,
				position: horizontal ? 'l' : 'b',
				majorGridlines: false,
				crossAxisId: secondId,
				auto: true,
				labelAlign: 'ctr',
				labelOffset: 100,
				noMultiLevelLabels: false,
			};
	const second: ChartAxis = {
		...common(),
		kind: 'val',
		id: secondId,
		position: horizontal ? 'b' : 'l',
		majorGridlines: true,
		crossAxisId: firstId,
		crossBetween: scatter ? 'midCat' : 'between',
	};
	return [first, second];
}

/** Fields only a category or date axis carries (CT_CatAx, CT_DateAx). */
const CATEGORY_ONLY = [
	'auto',
	'labelAlign',
	'labelOffset',
	'tickLabelSkip',
	'tickMarkSkip',
	'noMultiLevelLabels',
	'baseTimeUnit',
	'majorTimeUnit',
	'minorTimeUnit',
] as const;
/** Fields only a value axis carries (CT_ValAx). */
const VALUE_ONLY = ['crossBetween', 'displayUnits'] as const;

/**
 * `axis` as a `kind` axis: a category axis that becomes a value axis loses its label settings and
 * crosses between categories at their midpoints; a value axis that becomes a category axis loses
 * its scale bounds, units and display units and takes Excel's label defaults.
 */
export function convertChartAxis(axis: ChartAxis, kind: ChartAxisKind): ChartAxis {
	if (axis.kind === kind) return axis;
	const next: ChartAxis = structuredClone(axis);
	next.kind = kind;
	if (kind === 'val') {
		for (const key of CATEGORY_ONLY) delete next[key];
		next.crossBetween = 'midCat';
		return next;
	}
	for (const key of VALUE_ONLY) delete next[key];
	delete next.scaling.min;
	delete next.scaling.max;
	delete next.scaling.logBase;
	delete next.majorUnit;
	delete next.minorUnit;
	if (kind === 'cat') Object.assign(next, { auto: true, labelAlign: 'ctr', labelOffset: 100 });
	return next;
}

const SWAP: Readonly<Record<string, string>> = { b: 'l', l: 'b', t: 'r', r: 't' };

/** The position an axis takes when a chart turns between vertical and horizontal bars. */
export const rotatedAxisPosition = (position: string | undefined): string | undefined =>
	position === undefined ? undefined : (SWAP[position] ?? position);
