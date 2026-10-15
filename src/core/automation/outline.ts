import { createDocument, type Block } from '../docx/model';
import { ensureListDefinition } from '../docx/numbering-editing';
import { loadDocx } from '../docx/parse';
import { saveDocx } from '../docx/save';
import { createWorkbook, validateSheetName } from '../xlsx/workbook';
import { createEditSession } from '../xlsx/edit/session';
import { saveXlsx } from '../xlsx/write/save';

/**
 * Structured content a host (an assistant, a template, an import) turns into a new file. The
 * shapes are deliberately small: they name what Word and Excel already have a built-in style or
 * cell for, so the result is an ordinary document the editors open and change like any other.
 */
export type DocxOutlineBlock =
	| { type: 'title'; text: string }
	| { type: 'heading'; text: string; level?: 1 | 2 | 3 }
	| { type: 'paragraph'; text: string }
	| { type: 'bullets' | 'numbered'; items: readonly string[] };

export interface XlsxOutlineSheet {
	name: string;
	/** Row-major cell inputs; a string starting with `=` is a formula, as typed into a cell. */
	rows: readonly (readonly (string | number | boolean | null)[])[];
}

const MAX_BLOCKS = 2000;
const MAX_TEXT = 20_000;

function text(value: unknown): string {
	if (typeof value !== 'string' || value.length > MAX_TEXT)
		throw new Error('Outline text must be a string of at most 20000 characters');
	return value;
}

/** A new Word document from headings, paragraphs and lists, using the built-in styles. */
export async function createDocxFromOutline(
	outline: readonly DocxOutlineBlock[],
): Promise<Uint8Array> {
	if (!Array.isArray(outline) || outline.length > MAX_BLOCKS)
		throw new Error(`An outline has at most ${MAX_BLOCKS} blocks`);
	// Lists add numbering definitions, which only a loaded package can write, as in the editor.
	const loaded = await loadDocx(await saveDocx(createDocument()));
	const { model } = loaded;
	const blocks: Block[] = [];
	const id = () => `p${blocks.length + 1}`;
	for (const block of outline) {
		if (block.type === 'title' || block.type === 'heading' || block.type === 'paragraph') {
			const level = block.type === 'heading' ? (block.level ?? 1) : 0;
			if (![0, 1, 2, 3].includes(level)) throw new Error('Heading levels are 1, 2 or 3');
			const style =
				block.type === 'title' ? 'Title' : block.type === 'heading' ? `Heading${level}` : undefined;
			blocks.push({
				type: 'paragraph',
				id: id(),
				...(style ? { style } : {}),
				runs: [{ text: text(block.text) }],
			});
		} else if (block.type === 'bullets' || block.type === 'numbered') {
			if (!Array.isArray(block.items) || !block.items.length)
				throw new Error('A list needs at least one item');
			const list = ensureListDefinition(
				model.numberingCatalog,
				block.type === 'bullets' ? 'bullet' : 'decimal',
			);
			model.numberingCatalog = list.catalog;
			for (const item of block.items)
				blocks.push({
					type: 'paragraph',
					id: id(),
					numbering: { numId: list.numId, level: 0 },
					runs: [{ text: text(item) }],
				});
		} else throw new Error('Unknown outline block');
	}
	if (blocks.length) model.blocks = blocks;
	return loaded.save();
}

/** A new workbook with one sheet per entry, each filled from its rows starting at A1. */
export async function createXlsxFromSheets(
	sheets: readonly XlsxOutlineSheet[],
): Promise<Uint8Array> {
	if (!Array.isArray(sheets) || !sheets.length || sheets.length > 50)
		throw new Error('A workbook needs 1 to 50 sheets');
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	let cells = 0;
	sheets.forEach((sheet, index) => {
		const problem = validateSheetName(workbook, text(sheet.name), index === 0 ? 0 : -1);
		if (problem) throw new Error(problem);
		if (index === 0) session.renameSheet(0, sheet.name);
		else session.addSheet(sheet.name);
		const rows: XlsxOutlineSheet['rows'] = sheet.rows;
		// Hosts pass parsed JSON, so the shape is checked at run time too.
		if (!(rows instanceof Array)) throw new Error('Sheet rows must be an array');
		session.batch('Create from outline', () =>
			rows.forEach((row, r) => {
				if (!(row instanceof Array)) throw new Error('Each row must be an array');
				row.forEach((value, c) => {
					if (value === null || value === '') return;
					if (++cells > 100_000) throw new Error('A workbook outline has at most 100000 cells');
					if (!['string', 'number', 'boolean'].includes(typeof value))
						throw new Error('Cells hold text, numbers or booleans');
					session.setCellInput(
						index,
						r,
						c,
						typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : text(String(value)),
					);
				});
			}),
		);
	});
	return saveXlsx(workbook);
}
