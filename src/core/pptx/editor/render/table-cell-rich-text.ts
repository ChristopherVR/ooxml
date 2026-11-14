/**
 * table-cell-rich-text.ts: a table cell's text as rich segments, the same
 * `TextSegment` model shape text uses, so cell editing keeps (and can change)
 * per-run formatting instead of flattening the cell to one string.
 *
 * A loaded cell carries its runs as `textRuns` (sizes in points); an edited
 * cell gets `textSegments` (sizes in px, as in shape text), which the save
 * writes as `a:p`/`a:r` through the shape text writer while `text` still
 * spells them (`tableCellSegmentsForSave`). {@link withCellSegments} keeps
 * `text`, `textSegments`, `textRuns` and `paragraphs` in step so the renderer
 * (which paints `textRuns`) and the save agree.
 *
 * @module render/table-cell-rich-text
 */
import type {
	PptxTableCell,
	PptxTableCellParagraph,
	PptxTableCellStyle,
	PptxTableCellTextRun,
	TablePptxElement,
	TextSegment,
	TextStyle,
	XmlObject,
} from 'ooxml-core/pptx';
import { updateCellTextInRawXml } from 'ooxml-core/pptx';

import { tableCellSegmentsText } from '../../core/core/runtime/table-cell-segments-xml';
import { pairTableCellSource, tableCellSourceXml, withRunStyle } from './table-cell-source';

/** Shape text sizes are px; table runs and cell styles are points. */
const PX_PER_PT = 96 / 72;

const RUN_STYLE_KEYS = [
	'bold',
	'italic',
	'underline',
	'strikethrough',
	'color',
	'fontFamily',
	'baseline',
	'characterSpacing',
	'language',
] as const;

/** The run-level formatting of a cell run, as a shape text style. */
export function cellRunSegmentStyle(run: PptxTableCellTextRun): TextStyle {
	const style: TextStyle = {};
	for (const key of RUN_STYLE_KEYS) {
		const value = run[key];
		if (value !== undefined) {
			(style as Record<string, unknown>)[key] = value;
		}
	}
	if (run.fontSize !== undefined) {
		style.fontSize = run.fontSize * PX_PER_PT;
	}
	return style;
}

/** The text-level fields of a cell style, as a shape text run style. */
function cellStyleRunStyle(style: PptxTableCellStyle | undefined): TextStyle {
	const run: TextStyle = {};
	if (!style) {
		return run;
	}
	for (const key of ['bold', 'italic', 'underline', 'color', 'fontFamily'] as const) {
		if (style[key] !== undefined) {
			(run as Record<string, unknown>)[key] = style[key];
		}
	}
	if (style.fontSize !== undefined) {
		run.fontSize = style.fontSize * PX_PER_PT;
	}
	return run;
}

/** A segment's formatting as a cell run (sizes back in points). */
function segmentCellRun(segment: TextSegment): PptxTableCellTextRun {
	const run: PptxTableCellTextRun = { text: segment.text };
	const style = segment.style ?? {};
	for (const key of RUN_STYLE_KEYS) {
		const value = style[key];
		if (value !== undefined) {
			(run as unknown as Record<string, unknown>)[key] = value;
		}
	}
	if (typeof style.fontSize === 'number') {
		run.fontSize = Math.round((style.fontSize / PX_PER_PT) * 100) / 100;
	}
	return run;
}

/** A paragraph's own layout as the `paragraphProperties` the shape writer reads. */
function paragraphProperties(paragraph: PptxTableCellParagraph | undefined): TextStyle | undefined {
	if (!paragraph) {
		return undefined;
	}
	const { endParaFontSize: _endParaFontSize, ...layout } = paragraph;
	return Object.keys(layout).length > 0 ? (layout as TextStyle) : undefined;
}

function paragraphBreak(style: TextStyle): TextSegment {
	return { text: '\n', isParagraphBreak: true, style };
}

/**
 * A cell's text as segments: its own `textSegments` when they still spell its
 * text, else its loaded runs, else its flat text in the cell's run style.
 * Each paragraph's first segment carries that paragraph's layout. Given the
 * cell's source markup (`tableCellSourceXml`), each segment also carries the
 * run it was read from (`TextSegment.tableRun`) and each paragraph its `a:p`,
 * so saving the edit keeps the markup this model does not type.
 */
export function tableCellSegments(cell: PptxTableCell, cellXml?: XmlObject): TextSegment[] {
	if (cell.textSegments?.length && tableCellSegmentsText(cell.textSegments) === cell.text) {
		return cell.textSegments.map((segment) => ({ ...segment, style: { ...segment.style } }));
	}
	const source = cell.textRuns?.length ? pairTableCellSource(cell.textRuns, cellXml) : undefined;
	const segments: TextSegment[] = [];
	let paragraphIndex = 0;
	let paragraphStart = true;
	const push = (segment: TextSegment): void => {
		if (paragraphStart) {
			const layout = paragraphProperties(cell.paragraphs?.[paragraphIndex]);
			if (layout) {
				segment.paragraphProperties = layout;
			}
			const paragraph = source?.paragraphs[paragraphIndex];
			if (paragraph) {
				segment.tableParagraph = paragraph;
			}
			paragraphStart = false;
		}
		segments.push(segment);
	};
	const nextParagraph = (style: TextStyle): void => {
		if (paragraphStart) {
			push({ text: '', style: { ...style } });
		}
		segments.push(paragraphBreak({ ...style }));
		paragraphIndex++;
		paragraphStart = true;
	};
	const runs = cell.textRuns;
	if (runs?.length) {
		let lastStyle: TextStyle = cellStyleRunStyle(cell.style);
		runs.forEach((run, index) => {
			if (run.isParagraphBreak) {
				nextParagraph(lastStyle);
				return;
			}
			const runSource = source?.runs[index];
			if (run.isLineBreak) {
				const tableRun = withRunStyle(runSource, lastStyle);
				push({
					text: '\n',
					isLineBreak: true,
					style: { ...lastStyle },
					...(tableRun ? { tableRun } : {}),
				});
				return;
			}
			lastStyle = cellRunSegmentStyle(run);
			const segment: TextSegment = { text: run.text, style: lastStyle };
			const tableRun = withRunStyle(runSource, lastStyle);
			if (tableRun) {
				segment.tableRun = tableRun;
				if (tableRun.tag === 'a:fld') {
					// A field stays atomic: typing beside it never grows it.
					const type = tableRun.xml['@_type'];
					segment.fieldType = typeof type === 'string' && type ? type : 'field';
				}
			}
			push(segment);
		});
		if (paragraphStart) {
			push({ text: '', style: { ...lastStyle } });
		}
		return segments;
	}
	const base = cellStyleRunStyle(cell.style);
	(cell.text ?? '').split('\n').forEach((line, index) => {
		if (index > 0) {
			nextParagraph(base);
		}
		push({ text: line, style: { ...base } });
	});
	return segments;
}

/** The cell runs a renderer paints for these segments. */
export function cellRunsFromSegments(segments: readonly TextSegment[]): PptxTableCellTextRun[] {
	const runs: PptxTableCellTextRun[] = [];
	for (const segment of segments) {
		if (segment.isParagraphBreak) {
			runs.push({ text: '', isParagraphBreak: true });
		} else if (segment.isLineBreak) {
			runs.push({ text: '', isLineBreak: true });
		} else if (segment.text.length > 0) {
			runs.push(segmentCellRun(segment));
		}
	}
	return runs;
}

/**
 * Give a cell new rich content. `text` is derived from the segments exactly as
 * the save checks it, so the segments are what gets written; `textRuns` paint
 * them; `paragraphs` keep one layout entry per paragraph, taken from each
 * paragraph's first segment (falling back to the cell's existing entry).
 */
export function withCellSegments(cell: PptxTableCell, segments: TextSegment[]): PptxTableCell {
	const next: PptxTableCell = {
		...cell,
		text: tableCellSegmentsText(segments),
		textSegments: segments,
		textRuns: cellRunsFromSegments(segments),
	};
	// Renderers read an unset run flag as "inherit the cell's", so a format
	// turned off in the edit is turned off on the cell, where `false` wins over
	// a table style's bold header (runs set on still paint on).
	for (const key of ['bold', 'italic', 'underline'] as const) {
		if (segments.some((segment) => segment.text && segment.style?.[key] === false)) {
			next.style = { ...next.style, [key]: false };
		}
	}
	if (cell.paragraphs) {
		const paragraphs: PptxTableCellParagraph[] = [];
		let atStart = true;
		for (const segment of segments) {
			if (segment.isParagraphBreak) {
				atStart = true;
				continue;
			}
			if (atStart) {
				const own = segment.paragraphProperties as PptxTableCellParagraph | undefined;
				paragraphs.push({ ...(own ?? cell.paragraphs[paragraphs.length] ?? {}) });
				atStart = false;
			}
		}
		next.paragraphs = paragraphs;
	}
	return next;
}

/** Whether two segment lists spell and style the same content. */
export function sameCellSegments(a: readonly TextSegment[], b: readonly TextSegment[]): boolean {
	// The source markup a segment carries is not content: compare what it spells.
	const content = (segments: readonly TextSegment[]) =>
		JSON.stringify(segments, (key, value) =>
			key === 'tableRun' || key === 'tableParagraph' ? undefined : value,
		);
	return content(a) === content(b);
}

/**
 * Return a new table element with one cell's rich content replaced, or the
 * element itself when the cell does not exist or nothing changed (so leaving
 * a cell unedited writes nothing).
 */
export function setCellSegments(
	element: TablePptxElement,
	rowIndex: number,
	colIndex: number,
	segments: TextSegment[],
): TablePptxElement {
	const tableData = element.tableData;
	const cell = tableData?.rows[rowIndex]?.cells[colIndex];
	if (!tableData || !cell) {
		return element;
	}
	if (
		sameCellSegments(
			tableCellSegments(cell, tableCellSourceXml(element, rowIndex, colIndex)),
			segments,
		)
	) {
		return element;
	}
	const rows = tableData.rows.map((row, ri) =>
		ri !== rowIndex
			? row
			: {
					...row,
					cells: row.cells.map((item, ci) =>
						ci === colIndex ? withCellSegments(item, segments) : item,
					),
				},
	);
	return { ...element, tableData: { ...tableData, rows } };
}

/**
 * {@link setCellSegments} on an element, keeping `rawXml` in step with the
 * flat text so every reader of the markup sees the edit (the save rewrites
 * the cell's `a:txBody` from the segments).
 */
export function setTableElementCellSegments(
	element: TablePptxElement,
	rowIndex: number,
	colIndex: number,
	segments: TextSegment[],
): TablePptxElement {
	const updated = setCellSegments(element, rowIndex, colIndex, segments);
	if (updated === element || !element.rawXml) {
		return updated;
	}
	const text = updated.tableData?.rows[rowIndex]?.cells[colIndex]?.text ?? '';
	const rawXml = updateCellTextInRawXml(element, rowIndex, colIndex, text);
	return rawXml ? { ...updated, rawXml } : updated;
}
