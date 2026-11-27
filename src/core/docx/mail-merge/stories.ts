// Every list of blocks in a document, with the story it belongs to: the body, each header and
// footer, the footnotes and endnotes, the paragraphs of each table cell and the content of each
// text box. The headless merge reads and rewrites a document through these lists.
import type {
	Block,
	DocumentModel,
	HeaderFooterContent,
	HeaderFooterSlots,
	Paragraph,
	TableCell,
} from '../model';
import { cellSourceBlocks } from '../cell-blocks';

export type MergeStory = 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'textBox';

export interface BlockList {
	story: MergeStory;
	/** The list is a table cell's paragraphs: it can hold paragraphs only. */
	cell: boolean;
	/** The header or footer slot of the section (the first one, when slots share content). */
	slot?: HeaderFooterSlot;
}

export type HeaderFooterSlot = 'default' | 'first' | 'even';

/** New blocks for a list; the same array when nothing changed. */
export type ListMapper = (blocks: Block[], list: BlockList) => Block[];

const SLOTS: readonly HeaderFooterSlot[] = ['default', 'first', 'even'];

function mapList(blocks: Block[], list: BlockList, fn: ListMapper): Block[] {
	let changed = false;
	const inner = blocks.map((block): Block => {
		if (block.type === 'paragraph') {
			const next = mapTextBoxes(block, fn);
			changed ||= next !== block;
			return next;
		}
		let tableChanged = false;
		const rows = block.rows.map((row) =>
			row.map((cell): TableCell => {
				if (cell.sourceBlocks) {
					// A cell holding a nested table: its blocks in source order.
					const source = cellSourceBlocks(cell);
					const mapped = mapList(source, { ...list, cell: true }, fn);
					if (mapped === source) return cell;
					tableChanged = true;
					return {
						...cell,
						paragraphs: mapped.filter((item): item is Paragraph => item.type === 'paragraph'),
						sourceBlocks: mapped,
					};
				}
				const paragraphs = mapList(cell.paragraphs, { ...list, cell: true }, fn);
				if (paragraphs === cell.paragraphs) return cell;
				tableChanged = true;
				return { ...cell, paragraphs: paragraphs as Paragraph[] };
			}),
		);
		if (!tableChanged) return block;
		changed = true;
		return { ...block, rows };
	});
	const current = changed ? inner : blocks;
	return fn(current, list);
}

function mapTextBoxes(paragraph: Paragraph, fn: ListMapper): Paragraph {
	let changed = false;
	const runs = paragraph.runs.map((run) => {
		const content = run.image?.textBoxContent;
		// Editable text boxes are edited as their own story, as the editor's merge leaves them.
		if (!content?.blocks.length || run.image?.textBoxEditable) return run;
		const blocks = mapList(content.blocks, { story: 'textBox', cell: false }, fn);
		if (blocks === content.blocks) return run;
		changed = true;
		const image = { ...run.image!, textBoxContent: { ...content, blocks } };
		if (image.textBoxText)
			image.textBoxText = blocks.flatMap((block) =>
				block.type === 'paragraph' ? [block.runs.map((item) => item.text).join('')] : [],
			);
		return { ...run, image };
	});
	return changed ? { ...paragraph, runs } : paragraph;
}

/** `model` with `fn` applied to every block list, innermost first. */
export function mapBlockLists(model: DocumentModel, fn: ListMapper): DocumentModel {
	const blocks = mapList(model.blocks, { story: 'body', cell: false }, fn);
	const seen = new Map<HeaderFooterContent, HeaderFooterContent>();
	const slots = (item: HeaderFooterSlots | undefined, story: MergeStory) => {
		if (!item) return item;
		let changed = false;
		const next: HeaderFooterSlots = { ...item };
		for (const slot of SLOTS) {
			const content = item[slot];
			if (!content) continue;
			let mapped = seen.get(content);
			if (!mapped) {
				const blocks = mapList(content.blocks, { story, cell: false, slot }, fn);
				mapped = blocks === content.blocks ? content : { ...content, blocks };
				seen.set(content, mapped);
			}
			if (mapped === content) continue;
			changed = true;
			next[slot] = mapped;
		}
		return changed ? next : item;
	};
	const sections = model.sections?.map((section) => {
		const headers = slots(section.headers, 'header');
		const footers = slots(section.footers, 'footer');
		return headers === section.headers && footers === section.footers
			? section
			: { ...section, ...(headers && { headers }), ...(footers && { footers }) };
	});
	const notes = (items: NonNullable<DocumentModel['footnotes']>, story: MergeStory) =>
		items.map((note) => {
			const blocks = mapList(note.blocks, { story, cell: false }, fn);
			return blocks === note.blocks ? note : { ...note, blocks };
		});
	return {
		...model,
		blocks,
		...(sections && { sections }),
		...(model.footnotes && { footnotes: notes(model.footnotes, 'footnote') }),
		...(model.endnotes && { endnotes: notes(model.endnotes, 'endnote') }),
	};
}

/** Calls `visit` with every block list, without changing anything. */
export function forEachBlockList(
	model: DocumentModel,
	visit: (blocks: readonly Block[], list: BlockList) => void,
): void {
	mapBlockLists(model, (blocks, list) => {
		visit(blocks, list);
		return blocks;
	});
}
