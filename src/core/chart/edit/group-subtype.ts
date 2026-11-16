// Subtype settings of a chart group that are not ranged Series Options (`group-options.ts`): the
// 3-D bar shape and gap depth, the radar style, the surface wireframe flag, the pie-of-pie split and
// the bubble options. The target is the group that holds a series (flattened index, see
// `draft.ts`). Values follow the schema types of ECMA-376 Part 1, 21.2.3 (ST_Shape, ST_GapAmount for
// gap depth, ST_RadarStyle, ST_OfPieType, ST_SplitType, ST_SecondPieSize, ST_BubbleScale,
// ST_SizeRepresents); a setting the group element does not allow (`type-schema.ts`) is refused
// rather than written into a part Office would repair. `null` removes the element.
import type { ChartPlotGroup, ChartSpace } from '../model';
import { chartSeriesEntryAt, editChartSpace } from './draft';
import { GROUP_FIELDS } from './type-schema';

export type ChartBarShape = 'box' | 'cone' | 'coneToMax' | 'cylinder' | 'pyramid' | 'pyramidToMax';
export type ChartRadarStyle = 'standard' | 'marker' | 'filled';
export type ChartOfPieSplitType = 'auto' | 'cust' | 'percent' | 'pos' | 'val';

export interface ChartGroupSubtypeEdit {
	/** 3-D bar shape (`c:bar3DChart/c:shape`). */
	shape?: ChartBarShape | null;
	/** Depth gap of a 3-D bar, line or area group, percent (ST_GapAmount 0..500). */
	gapDepth?: number | null;
	radarStyle?: ChartRadarStyle | null;
	/** Surface: drawn as a wireframe. */
	wireframe?: boolean | null;
	ofPieType?: 'pie' | 'bar' | null;
	splitType?: ChartOfPieSplitType | null;
	/** The split threshold (`c:splitPos`). */
	splitPosition?: number | null;
	/** Point indexes drawn in the second plot of a custom split. */
	customSplit?: number[] | null;
	/** Second plot size, percent of the first (ST_SecondPieSize 5..200). */
	secondPieSize?: number | null;
	/** Pie-of-pie series lines (`c:serLines`). */
	seriesLines?: boolean | null;
	bubble3D?: boolean | null;
	/** Bubble size scale, percent (ST_BubbleScale 0..300). */
	bubbleScale?: number | null;
	showNegativeBubbles?: boolean | null;
	sizeRepresents?: 'area' | 'w' | null;
}

const SHAPES: readonly string[] = [
	'box',
	'cone',
	'coneToMax',
	'cylinder',
	'pyramid',
	'pyramidToMax',
];
const CHOICES: Partial<Record<keyof ChartGroupSubtypeEdit, readonly string[]>> = {
	shape: SHAPES,
	radarStyle: ['standard', 'marker', 'filled'],
	ofPieType: ['pie', 'bar'],
	splitType: ['auto', 'cust', 'percent', 'pos', 'val'],
	sizeRepresents: ['area', 'w'],
};
const RANGES: Partial<Record<keyof ChartGroupSubtypeEdit, readonly [number, number, boolean]>> = {
	gapDepth: [0, 500, true],
	secondPieSize: [5, 200, true],
	bubbleScale: [0, 300, true],
	splitPosition: [-Infinity, Infinity, false],
};
const KEYS: readonly (keyof ChartGroupSubtypeEdit)[] = [
	'shape',
	'gapDepth',
	'radarStyle',
	'wireframe',
	'ofPieType',
	'splitType',
	'splitPosition',
	'customSplit',
	'secondPieSize',
	'seriesLines',
	'bubble3D',
	'bubbleScale',
	'showNegativeBubbles',
	'sizeRepresents',
];

function check(key: keyof ChartGroupSubtypeEdit, value: unknown): void {
	const choices = CHOICES[key];
	if (choices && !choices.includes(value as string))
		throw new RangeError(`Invalid chart ${key}: ${String(value)}`);
	const range = RANGES[key];
	if (range) {
		const [min, max, whole] = range;
		if (
			typeof value !== 'number' ||
			!Number.isFinite(value) ||
			(whole && !Number.isInteger(value)) ||
			value < min ||
			value > max
		)
			throw new RangeError(`Chart ${key} must be a number from ${min} to ${max}`);
	}
	if (
		key === 'customSplit' &&
		(!Array.isArray(value) || value.some((index) => !Number.isInteger(index) || index < 0))
	)
		throw new RangeError('A custom split lists whole point indexes');
	if (
		['wireframe', 'seriesLines', 'bubble3D', 'showNegativeBubbles'].includes(key) &&
		typeof value !== 'boolean'
	)
		throw new RangeError(`Chart ${key} is a flag`);
}

/** Edits the subtype settings of the group holding series `seriesIndex`. */
export function setChartGroupSubtype(
	space: ChartSpace,
	seriesIndex: number,
	edit: ChartGroupSubtypeEdit,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { group } = chartSeriesEntryAt(draft, seriesIndex);
		const allowed = GROUP_FIELDS[group.element] ?? [];
		const record = group as unknown as Record<string, unknown>;
		for (const key of KEYS) {
			const value = edit[key];
			if (value === undefined) continue;
			if (!allowed.includes(key as keyof ChartPlotGroup))
				throw new RangeError(`A ${group.element} has no ${key} setting`);
			if (value === null) {
				delete record[key];
				if (key === 'seriesLines') delete group.seriesLinesSpPr;
				continue;
			}
			check(key, value);
			record[key] = Array.isArray(value) ? [...value] : value;
		}
	});
}
