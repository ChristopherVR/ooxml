// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Where the fill-in fields of a document are: legacy form fields (a begin marker carrying
// `w:ffData`) and content controls of the kinds a user fills (text, check box, list, date),
// around runs, around a table cell or around paragraphs. Found in reading order, the body first,
// then each header and footer part once.
import type { ContentControl, ContentControlKind } from './content-control-model';
import type { FormField } from './form-field-model';
import type { Block, DocumentModel, Paragraph, Table, TableCell, TextRun } from './model';
import { resolveRunFormatting } from './run-formatting';
import { shownText } from './style-ref';

/** The story a field is in. */
export type FormFieldStory = 'body' | 'header' | 'footer';

/**
 * The text a run of a paragraph shows, as a person reads the form: none for hidden text, set on
 * the run or through its paragraph or character style, or for text a revision deletes or moves.
 */
export type ShownText = (run: TextRun, paragraph: Paragraph) => string;

/** How the runs of `model` show, through its styles. */
const shownTextIn =
	(model: DocumentModel): ShownText =>
	(run, paragraph) =>
		resolveRunFormatting(run, {
			runCatalog: model.characterStyles,
			paragraphCatalog: model.paragraphStyles,
			paragraphStyleId: paragraph.style,
		}).vanish
			? ''
			: shownText(run);

/** A cell's place in its table: its row, its index in the row and the grid columns it covers. */
export interface CellPlace {
	table: Table;
	row: number;
	cell: number;
	column: number;
	span: number;
	/** The cell's paragraph holding the field. */
	paragraph: number;
}

/** A fill-in field found in the model, with references into it for reading and writing. */
export interface FieldSite {
	/** A legacy field, a control around runs, a control around a cell, or one around paragraphs. */
	kind: 'legacy' | 'inline' | 'cell' | 'block';
	story: FormFieldStory;
	/** The story's blocks, and the index of the block holding the field (or its first block). */
	blocks: Block[];
	block: number;
	/** The paragraph holding a legacy field or an inline control; a cell or block control's first. */
	paragraph: Paragraph;
	/** A legacy field's runs from its begin to its end marker; an inline control's runs. */
	from: number;
	to: number;
	/** Set when the field is in a table cell, or is a control around one. */
	place?: CellPlace;
	/** A block control's paragraphs, as indexes into `blocks`. */
	paragraphs?: number[];
	formField?: FormField;
	control?: ContentControl;
	/** The field is in a table nested in a cell, which this library keeps but does not edit. */
	nested?: boolean;
	/** The text the runs of its document show, for its value and the labels around it. */
	shown: ShownText;
}

/** The content control kinds a user fills in. */
const FILLED: ReadonlySet<ContentControlKind> = new Set([
	'richText',
	'text',
	'checkbox',
	'dropDownList',
	'comboBox',
	'date',
]);

export const isFilledControl = (control: ContentControl | undefined): control is ContentControl =>
	Boolean(control && FILLED.has(control.kind));

/** The index of the marker that ends the complex field begun at `begin`. */
export function fieldEnd(runs: readonly TextRun[], begin: number): number | undefined {
	let depth = 0;
	for (let index = begin; index < runs.length; index++) {
		if (runs[index]!.fieldChar === 'begin') depth++;
		else if (runs[index]!.fieldChar === 'end' && --depth === 0) return index;
	}
	return undefined;
}

/** Whether a run is more than text: a field marker or code, a picture, an equation, a note. */
const isObject = (run: TextRun) =>
	Boolean(run.fieldChar || run.fieldCode !== undefined || run.image || run.equation) ||
	Boolean(run.noteReference || run.noteMark || run.noteSeparator);

type Base = Omit<FieldSite, 'kind' | 'paragraph' | 'from' | 'to'>;

/** The legacy fields and inline controls of one paragraph, by where they start. */
function paragraphSites(paragraph: Paragraph, base: Base): FieldSite[] {
	const runs = paragraph.runs;
	const sites: FieldSite[] = [];
	runs.forEach((run, index) => {
		if (run.fieldChar === 'begin' && run.formField)
			sites.push({
				...base,
				kind: 'legacy',
				paragraph,
				from: index,
				to: fieldEnd(runs, index) ?? index,
				formField: run.formField,
			});
		const control = run.contentControl;
		if (control?.level !== 'inline' || runs[index - 1]?.contentControl?.id === control.id) return;
		let last = index;
		while (runs[last + 1]?.contentControl?.id === control.id) last++;
		// A control that holds fields, pictures or other objects is not filled as text.
		const content = runs.slice(index, last + 1);
		if (!isFilledControl(control) || (control.kind !== 'checkbox' && content.some(isObject)))
			return;
		sites.push({ ...base, kind: 'inline', paragraph, from: index, to: last, control });
	});
	return sites;
}

/** Whether a paragraph holds a fill-in field of its own. */
export const holdsField = (paragraph: Paragraph): boolean =>
	isFilledControl(paragraph.contentControl) ||
	paragraphSites(paragraph, { story: 'body', blocks: [], block: 0, shown: shownText }).length > 0;

function tableSites(table: Table, base: Base, nested: boolean): FieldSite[] {
	const sites: FieldSite[] = [];
	table.rows.forEach((cells, row) => {
		let column = 0;
		cells.forEach((cell, index) => {
			const span = cell.gridSpan ?? 1;
			const place = (paragraph: number): CellPlace => ({
				table,
				row,
				cell: index,
				column,
				span,
				paragraph,
			});
			const inner = cell.paragraphs.flatMap((paragraph, at) =>
				paragraphSites(paragraph, { ...base, place: place(at), ...(nested && { nested }) }),
			);
			const first = cell.paragraphs[0];
			if (isFilledControl(cell.contentControl) && first && !inner.length)
				sites.push({
					...base,
					kind: 'cell',
					paragraph: first,
					from: 0,
					to: first.runs.length - 1,
					place: place(0),
					control: cell.contentControl,
					...(nested && { nested }),
				});
			sites.push(...inner);
			for (const block of cell.sourceBlocks ?? [])
				if (block.type === 'table') sites.push(...tableSites(block, base, true));
			column += span;
		});
	});
	return sites;
}

/** The fill-in fields of one story, in reading order. */
function storySites(blocks: Block[], story: FormFieldStory, shown: ShownText): FieldSite[] {
	const sites: FieldSite[] = [];
	for (let index = 0; index < blocks.length; index++) {
		const block = blocks[index]!;
		if (block.type === 'table') {
			sites.push(...tableSites(block, { story, blocks, block: index, shown }, false));
			continue;
		}
		const control = block.contentControl;
		if (control?.level === 'block' && blocks[index - 1]?.contentControl?.id !== control.id) {
			const members: number[] = [];
			for (let at = index; blocks[at]?.contentControl?.id === control.id; at++) members.push(at);
			// A control around a table, or around fields of its own, is not filled as text.
			const container = members.some((at) => {
				const item = blocks[at]!;
				return (
					item.type === 'table' ||
					paragraphSites(item, { story, blocks, block: at, shown }).length > 0
				);
			});
			if (isFilledControl(control) && !container && control.kind !== 'checkbox')
				sites.push({
					kind: 'block',
					story,
					blocks,
					block: index,
					paragraph: block,
					from: 0,
					to: block.runs.length - 1,
					paragraphs: members,
					control,
					shown,
				});
		}
		sites.push(...paragraphSites(block, { story, blocks, block: index, shown }));
	}
	return sites;
}

/** Every fill-in field of the document: the body, then each header and footer part once. */
export function formFieldSites(model: DocumentModel): FieldSite[] {
	const shown = shownTextIn(model);
	const sites = storySites(model.blocks, 'body', shown);
	const seen = new Set<Block[]>();
	for (const section of model.sections ?? [])
		for (const [story, slots] of [
			['header', section.headers],
			['footer', section.footers],
		] as const)
			for (const content of [slots?.default, slots?.first, slots?.even]) {
				if (!content || seen.has(content.blocks)) continue;
				seen.add(content.blocks);
				sites.push(...storySites(content.blocks, story, shown));
			}
	return sites;
}

/** The runs between a legacy field's `separate` and `end` markers: its result. */
export function legacyResult(site: FieldSite): { from: number; to: number } | undefined {
	const runs = site.paragraph.runs;
	let depth = 0;
	for (let index = site.from; index <= site.to; index++) {
		const run = runs[index]!;
		if (run.fieldChar === 'begin') depth++;
		else if (run.fieldChar === 'end') depth--;
		else if (run.fieldChar === 'separate' && depth === 1)
			return { from: index + 1, to: site.to - 1 };
	}
	return undefined;
}

/** The cell a site is in. */
export const siteCell = (site: FieldSite): TableCell | undefined =>
	site.place && site.place.table.rows[site.place.row]?.[site.place.cell];
