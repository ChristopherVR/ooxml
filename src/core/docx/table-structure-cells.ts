// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Cell construction helpers shared by the table structure edits.
import type { Paragraph, Table, TableCell } from './model';
import type { Twips } from './units';

/** Hands out paragraph ids that no paragraph in the document has. */
export type IdFactory = () => string;

const PARAGRAPH_FORMAT_KEYS = [
	'style',
	'align',
	'justification',
	'direction',
	'spacingBeforeTwips',
	'spacingAfterTwips',
	'lineSpacingTwips',
	'lineSpacingRule',
	'indentLeftTwips',
	'indentRightTwips',
	'indentStartTwips',
	'indentEndTwips',
	'firstLineTwips',
	'hangingTwips',
] as const;

/** An empty paragraph that takes the paragraph formatting (not the text) of `like`. */
export function emptyParagraph(newId: IdFactory, like?: Paragraph): Paragraph {
	const paragraph: Paragraph = { type: 'paragraph', id: newId(), runs: [{ text: '' }] };
	if (like) {
		const source = like as unknown as Record<string, unknown>;
		const target = paragraph as unknown as Record<string, unknown>;
		for (const key of PARAGRAPH_FORMAT_KEYS)
			if (source[key] !== undefined) target[key] = source[key];
	}
	return paragraph;
}

/** A new empty cell with the formatting of `like`, the given span and width, and a fresh paragraph. */
export function emptyCell(
	newId: IdFactory,
	like: TableCell | undefined,
	span: number,
	width: Twips | undefined,
): TableCell {
	const cell: TableCell = { paragraphs: [emptyParagraph(newId, like?.paragraphs[0])] };
	if (span > 1) cell.gridSpan = span;
	if (width !== undefined) cell.widthTwips = width;
	if (like) {
		if (like.verticalAlign) cell.verticalAlign = like.verticalAlign;
		if (like.shadingFill) cell.shadingFill = like.shadingFill;
		if (like.shadingThemeFill) cell.shadingThemeFill = structuredClone(like.shadingThemeFill);
		if (like.borders) cell.borders = structuredClone(like.borders);
		if (like.margins) cell.margins = structuredClone(like.margins);
	}
	return cell;
}

/** A deep copy of the table whose structure an edit may change freely. */
export function copyTable(table: Table): Table {
	return structuredClone(table);
}

/** The row properties of a new row that copies `from`: height, keep-together and header flags. */
export function copiedRowProperties(
	table: Table,
	from: number,
): NonNullable<Table['rowProperties']>[number] {
	const { contentControl: _control, ...properties } = table.rowProperties?.[from] ?? {};
	return structuredClone(properties);
}

type RowProperties = NonNullable<Table['rowProperties']>;

/** The table's row properties padded to one entry per row. */
export function paddedRowProperties(table: Table): RowProperties {
	return Array.from(
		{ length: table.rows.length },
		(_, index) => table.rowProperties?.[index] ?? {},
	);
}

/** Stores row properties, dropping the list when no row has any. */
export function storeRowProperties(table: Table, properties: RowProperties): void {
	if (properties.some((entry) => Object.keys(entry).length)) table.rowProperties = properties;
	else delete table.rowProperties;
}
