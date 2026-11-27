/** The pictures, charts and shapes a laid-out page draws, for the export to load before painting. */
import type { LayoutBlockBox, LayoutPageBox } from '../layout/index';
import type { DocxPdfPicture } from './pdf-pictures';

export function* blockPictures(blocks: readonly LayoutBlockBox[]): Generator<DocxPdfPicture> {
	for (const block of blocks) {
		if (block.kind === 'paragraph') {
			for (const line of block.lines)
				for (const fragment of line.fragments) if (fragment.object) yield fragment.object;
		} else
			for (const row of block.rows) {
				for (const cell of row.cells) yield* blockPictures(cell);
				for (const nested of row.cellTables ?? []) yield* blockPictures(nested);
			}
	}
}

export function* pagePictures(page: LayoutPageBox): Generator<DocxPdfPicture> {
	for (const column of page.columns) yield* blockPictures(column.blocks);
	for (const note of page.footnotes ?? []) yield* blockPictures(note.paragraphs);
	yield* page.floats ?? [];
}
