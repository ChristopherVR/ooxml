import type { ChartSpace } from '../model';
import { editChartSpace, assignOrDelete } from './draft';
import { editChartTextBody, type ChartTextEdit } from './text';

export interface ChartLegendEntryEdit {
	index: number;
	deleted?: boolean;
	text?: ChartTextEdit;
}

/** Replace the entry override list while retaining the rich text of surviving entries. */
export function setChartLegendEntries(
	space: ChartSpace,
	entries: readonly ChartLegendEntryEdit[],
): ChartSpace {
	const indices = new Set<number>();
	for (const entry of entries) {
		if (!Number.isSafeInteger(entry.index) || entry.index < 0 || indices.has(entry.index))
			throw new RangeError('Legend entry indices must be unique non-negative integers');
		indices.add(entry.index);
	}
	return editChartSpace(space, (draft) => {
		if (!draft.legend) throw new RangeError('The chart has no legend');
		const old = draft.legend.entries;
		draft.legend.entries = entries.map((edit) => {
			const entry = { ...old.find((item) => item.index === edit.index), index: edit.index };
			assignOrDelete(entry, 'deleted', edit.deleted);
			if (edit.text) entry.txPr = editChartTextBody(entry.txPr, edit.text);
			return entry;
		});
	});
}
