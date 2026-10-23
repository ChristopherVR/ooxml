// Series Options of a chart group: bar gap width and overlap, the first slice angle of a pie or
// doughnut and a doughnut's hole size. A Format pane edits them on a series, so the target is the
// group that holds the series (flattened index, see `draft.ts`). Ranges follow the schema types
// PowerPoint's and Excel's panes enforce (ST_GapAmount 0..500, ST_Overlap -100..100,
// ST_FirstSliceAng 0..360, the panes' 10..90 for ST_HoleSize); a setting the group element does not
// allow (`type-schema.ts`) is refused rather than written into a part Office would repair.
import type { ChartSpace } from '../model';
import { chartSeriesEntryAt, editChartSpace } from './draft';
import { GROUP_FIELDS } from './type-schema';

export interface ChartGroupOptionsEdit {
	/** Gap between categories, percent of a bar's width (bar, 3-D bar, bar of pie). */
	gapWidth?: number;
	/** Overlap of a category's bars, percent (2-D bar). */
	overlap?: number;
	/** Angle of the first slice, degrees (pie, doughnut). */
	firstSliceAngle?: number;
	/** Doughnut hole size, percent of the diameter. */
	holeSize?: number;
}

const RANGES: Record<keyof ChartGroupOptionsEdit, readonly [number, number]> = {
	gapWidth: [0, 500],
	overlap: [-100, 100],
	firstSliceAngle: [0, 360],
	holeSize: [10, 90],
};

/** Edits the options of the group holding series `seriesIndex`. */
export function setChartGroupOptions(
	space: ChartSpace,
	seriesIndex: number,
	edit: ChartGroupOptionsEdit,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { group } = chartSeriesEntryAt(draft, seriesIndex);
		const allowed = GROUP_FIELDS[group.element] ?? [];
		for (const key of Object.keys(RANGES) as (keyof ChartGroupOptionsEdit)[]) {
			const value = edit[key];
			if (value === undefined) continue;
			const [min, max] = RANGES[key];
			if (!Number.isInteger(value) || value < min || value > max)
				throw new RangeError(`Chart ${key} must be a whole number from ${min} to ${max}`);
			if (!allowed.includes(key)) throw new RangeError(`A ${group.element} has no ${key} setting`);
			group[key] = value;
		}
	});
}
