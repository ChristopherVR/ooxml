import type { DocumentModel } from 'ooxml-core/docx';
import { layoutDocumentModel } from 'ooxml-core/docx/layout';
import type { ReferencePage } from 'ooxml-core/docx/ui';
import { createCanvasMeasurer } from '../canvas-measurer';
import { pageNumbers } from '../print-header-footer';

type PagePart = { number: string; end: number };

export function referencePages(model: DocumentModel): ReferencePage {
	const { pages } = layoutDocumentModel(model, createCanvasMeasurer());
	const numbers = pageNumbers(model, pages);
	const fragments = new Map<string, PagePart[]>();
	const add = (id: string, number: string, end: number) => {
		const parts = fragments.get(id) ?? [];
		parts.push({ number, end });
		fragments.set(id, parts);
	};
	for (const [index, page] of pages.entries()) {
		const number = numbers[index];
		if (number === undefined) continue;
		for (const column of page.columns)
			for (const block of column.blocks) {
				if (block.kind === 'paragraph')
					add(block.blockId, number, block.lines.at(-1)?.sourceEnd ?? Infinity);
				else
					for (const row of block.rows) {
						if (row.repeated) continue;
						for (const cell of row.cells)
							for (const box of cell)
								add(box.blockId, number, box.lines.at(-1)?.sourceEnd ?? Infinity);
					}
			}
	}
	return (id, offset, end) => {
		const parts = fragments.get(id);
		if (end || offset === undefined) return (end ? parts?.at(-1) : parts?.[0])?.number;
		return (parts?.find((part) => offset < part.end) ?? parts?.at(-1))?.number;
	};
}
