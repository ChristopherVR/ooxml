/**
 * The renderer view (`PptxTableCell.textRuns`) of a new cell's rich segments,
 * so a table made in code draws its runs before it is saved and reloaded.
 *
 * @module sdk/table-cell-runs-input
 */

import type { PptxTableCellTextRun } from '../../types/table';
import type { TextSegment } from '../../types/text';

/** Segments (font size in px) → table cell runs (font size in points). */
export function tableCellRunsFromSegments(
	segments: readonly TextSegment[],
): PptxTableCellTextRun[] {
	const runs: PptxTableCellTextRun[] = [];
	const pushText = (text: string, style: TextSegment['style']): void => {
		const run: PptxTableCellTextRun = { text };
		if (style?.bold !== undefined) {
			run.bold = style.bold;
		}
		if (style?.italic !== undefined) {
			run.italic = style.italic;
		}
		if (style?.underline !== undefined) {
			run.underline = style.underline;
		}
		if (style?.strikethrough !== undefined) {
			run.strikethrough = style.strikethrough;
		}
		if (style?.color !== undefined) {
			run.color = style.color;
		}
		if (style?.fontSize !== undefined) {
			run.fontSize = (style.fontSize * 72) / 96;
		}
		if (style?.fontFamily !== undefined) {
			run.fontFamily = style.fontFamily;
		}
		if (style?.characterSpacing !== undefined) {
			run.characterSpacing = style.characterSpacing;
		}
		runs.push(run);
	};
	for (const segment of segments) {
		if (segment.isParagraphBreak) {
			runs.push({ text: '', isParagraphBreak: true });
			continue;
		}
		if (segment.isLineBreak) {
			runs.push({ text: '', isLineBreak: true });
			continue;
		}
		segment.text.split('\n').forEach((part, index) => {
			if (index > 0) {
				runs.push({ text: '', isParagraphBreak: true });
			}
			if (part) {
				pushText(part, segment.style);
			}
		});
	}
	return runs;
}
