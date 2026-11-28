// Replacing every content control of a document by its content, before an export to PDF: the
// content keeps its formatting, and a control showing its placeholder leaves its placeholder text
// as Word does. Optionally the paragraphs that are hidden whole (their paragraph mark and every
// run hidden text) go too.
import { mapBlockLists } from '../mail-merge/stories';
import type { Block, DocumentModel, Paragraph, Table, TableCell, TextRun } from '../model';
import { paragraphMarkRun } from '../paragraph-mark';
import { resolveRunFormatting } from '../run-formatting';
import { withChain } from './chains';
import type { FlattenContentControlsOptions } from './types';

const plainRun = (run: TextRun): TextRun => {
	if (!run.contentControl) return run;
	const { contentControl: _control, outerContentControls: _outer, ...rest } = run;
	return rest;
};

/**
 * `holder` without its controls. It says it has none around it either (an empty
 * `outerContentControls`), so that the writer unwraps the outer ones too.
 */
function withoutControls<T extends Parameters<typeof withChain>[0]>(holder: T): T {
	if (!holder.contentControl) return holder;
	return { ...withChain(holder, []), outerContentControls: [] };
}

const plainParagraph = (paragraph: Paragraph): Paragraph =>
	withoutControls({ ...paragraph, runs: paragraph.runs.map(plainRun) });

function plainCell(cell: TableCell): TableCell {
	const { contentControl: _control, ...rest } = cell;
	const paragraphs = cell.paragraphs.map(plainParagraph);
	if (!cell.sourceBlocks) return { ...rest, paragraphs };
	const byId = new Map(paragraphs.map((paragraph) => [paragraph.id, paragraph]));
	return {
		...rest,
		paragraphs,
		sourceBlocks: cell.sourceBlocks.map((block) =>
			block.type === 'paragraph' ? (byId.get(block.id) ?? block) : plainTable(block),
		),
	};
}

function plainTable(table: Table): Table {
	return withoutControls({
		...table,
		rows: table.rows.map((row) => row.map(plainCell)),
		...(table.rowProperties && {
			rowProperties: table.rowProperties.map((properties) => withoutControls(properties)),
		}),
	});
}

/** `model` without content controls (see above). */
export function flattenContentControls(
	model: DocumentModel,
	options: FlattenContentControlsOptions = {},
): DocumentModel {
	const hidden = (paragraph: Paragraph) => {
		const context = {
			runCatalog: model.characterStyles,
			paragraphCatalog: model.paragraphStyles,
			paragraphStyleId: paragraph.style,
		};
		const vanish = (run: TextRun) => Boolean(resolveRunFormatting(run, context).vanish);
		return (
			vanish(paragraphMarkRun(paragraph)) && paragraph.runs.every((run) => !run.text || vanish(run))
		);
	};
	return mapBlockLists(model, (blocks, list) => {
		if (list.story === 'textBox') return blocks;
		let output: Block[] = blocks.map((block) =>
			block.type === 'paragraph' ? plainParagraph(block) : list.cell ? block : plainTable(block),
		);
		if (options.removeHiddenParagraphs) {
			const kept = output.filter((block) => block.type !== 'paragraph' || !hidden(block));
			// A cell or story keeps one paragraph.
			if (kept.some((block) => block.type === 'paragraph')) output = kept;
		}
		return output;
	});
}
