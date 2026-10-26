import type { ChartSpace } from '../model';
import { editChartSpace } from './draft';
import { GROUP_FIELDS } from './type-schema';
import { editChartShapeProperties, type ChartShapeEdit } from './shape-properties';

export interface ChartGroupDecorationsEdit {
	dropLines?: ChartShapeEdit | null;
	hiLowLines?: ChartShapeEdit | null;
	seriesLines?: ChartShapeEdit | null;
	upDownBars?: { gapWidth?: number | null; up?: ChartShapeEdit; down?: ChartShapeEdit } | null;
}

/** Apply only to groups whose schema permits the decoration, retaining source shape content. */
export function setChartGroupDecorations(
	space: ChartSpace,
	edit: ChartGroupDecorationsEdit,
): ChartSpace {
	if (
		edit.upDownBars?.gapWidth != null &&
		(!Number.isInteger(edit.upDownBars.gapWidth) ||
			edit.upDownBars.gapWidth < 0 ||
			edit.upDownBars.gapWidth > 500)
	)
		throw new RangeError('Invalid up/down bar gap width');
	return editChartSpace(space, (draft) => {
		for (const group of draft.plotArea.groups) {
			const allowed = GROUP_FIELDS[group.element] ?? [];
			for (const key of ['dropLines', 'hiLowLines', 'seriesLines'] as const) {
				const value = edit[key];
				if (value === undefined || !allowed.includes(key)) continue;
				const shape = `${key}SpPr` as const;
				if (value === null) {
					delete group[key];
					delete group[shape];
				} else {
					group[key] = true;
					group[shape] = editChartShapeProperties(group[shape], value);
				}
			}
			if (edit.upDownBars === undefined || !allowed.includes('upDownBars')) continue;
			if (edit.upDownBars === null) {
				delete group.upDownBars;
				delete group.upDownBarsGapWidth;
				delete group.upBars;
				delete group.downBars;
			} else {
				group.upDownBars = true;
				group.upBars ??= {};
				group.downBars ??= {};
				const { gapWidth, up, down } = edit.upDownBars;
				if (gapWidth === null) delete group.upDownBarsGapWidth;
				else if (gapWidth !== undefined) group.upDownBarsGapWidth = gapWidth;
				if (up) {
					group.upBars ??= {};
					group.upBars.spPr = editChartShapeProperties(group.upBars.spPr, up);
				}
				if (down) {
					group.downBars ??= {};
					group.downBars.spPr = editChartShapeProperties(group.downBars.spPr, down);
				}
			}
		}
	});
}
