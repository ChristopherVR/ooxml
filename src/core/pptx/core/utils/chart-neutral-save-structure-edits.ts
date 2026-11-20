// The rows and series a chart's data sheet added or removed (`PptxChartData.structureEdits`),
// replayed through the data sheet operations Word, Excel and Visio apply. Nothing here infers a
// row or series from the difference between two states: a blank row inside one series' view and
// a row past the end of another's read the same, so a difference cannot tell them apart.
import { applyChartSpaceDataSheetEdits, chartSeriesEntries } from '../../../chart/index';
import type { PptxChartData } from '../types';
import { chartDataApplySheetEdits } from './chart-data-sheet';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

/**
 * Claims the structure edits of `probe` (the chart as edited) and returns `before` as they leave
 * it, so the planners after this one see only cell edits.
 */
export function planNeutralStructureEdits(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): PptxChartData {
	const replayed = probe.structureEdits;
	delete probe.structureEdits;
	if (!replayed?.length) return before;
	const { structureEdits: _recorded, ...aligned } =
		chartDataApplySheetEdits(before, replayed) ?? before;
	const seriesCount = aligned.series.length;
	edits.push((space) => {
		const next = applyChartSpaceDataSheetEdits(space, replayed);
		// The part refused an edit the chart's data took: the object-tree writers save it instead.
		if (chartSeriesEntries(next).length !== seriesCount)
			throw new RangeError('The chart part does not take the recorded data sheet edits');
		return next;
	});
	return aligned;
}
