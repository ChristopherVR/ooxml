// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Labels for the fill-in fields of a form, read from the layout around each field as a person
// reads a form: in a grid, the nearest label cell to its left, with the row's first label; else
// the column's header; the text beside the field in its own paragraph (after a check box, the text
// after it); in the body, the paragraph before it. Repeated labels are numbered ("To 1", "To 2").
import { isFilledControl, holdsField, siteCell, type FieldSite } from './form-fill-sites';
import type { Paragraph, TableCell, TextRun } from './model';
import { shownText } from './style-ref';

/** Text as a label: one line, single spaces, without a trailing colon. */
export const labelText = (text: string): string => {
	const line = text.replace(/\s+/gu, ' ').trim();
	return line.endsWith(':') ? line.slice(0, -1).trimEnd() : line;
};

const runsText = (runs: readonly TextRun[]) =>
	labelText(
		runs
			.filter((run) => !run.fieldChar && run.fieldCode === undefined)
			.map(shownText)
			.join(''),
	);
const paragraphText = (paragraph: Paragraph) => runsText(paragraph.runs);

/** A cell that labels others: it has text and no fill-in field. */
function cellLabel(cell: TableCell | undefined): string | undefined {
	if (!cell || isFilledControl(cell.contentControl) || cell.paragraphs.some(holdsField))
		return undefined;
	return labelText(cell.paragraphs.map(paragraphText).join(' ')) || undefined;
}

/** The grid columns each cell of a row covers. */
function columns(cells: readonly TableCell[]) {
	let column = 0;
	return cells.map((cell) => {
		const span = cell.gridSpan ?? 1;
		const placed = { cell, column, span };
		column += span;
		return placed;
	});
}

/** A field's labels from its row: the nearest label cell to the left, after the row's first. */
function rowLabels(site: FieldSite): string[] {
	const place = site.place!;
	const cells = place.table.rows[place.row] ?? [];
	for (let index = place.cell - 1; index >= 0; index--) {
		const left = cellLabel(cells[index]);
		if (!left) continue;
		const first = index > 0 ? cellLabel(cells[0]) : undefined;
		return first && first !== left ? [first, left] : [left];
	}
	return [];
}

/** The nearest label cell above the field's cell that covers all of its grid columns. */
function columnHeader(site: FieldSite): string | undefined {
	const { table, row, column, span } = site.place!;
	for (let index = row - 1; index >= 0; index--) {
		const above = columns(table.rows[index] ?? []).find(
			(item) => item.column <= column && item.column + item.span >= column + span,
		);
		const text = cellLabel(above?.cell);
		if (text) return text;
	}
	return undefined;
}

/** The text beside a field in its own paragraph, up to the fields next to it. */
function besideText(site: FieldSite, siblings: readonly FieldSite[]): string | undefined {
	if (site.kind === 'cell' || site.kind === 'block') return undefined;
	const mine = siblings.filter((item) => item.paragraph === site.paragraph);
	const at = mine.indexOf(site);
	const runs = site.paragraph.runs;
	const before = runsText(runs.slice((mine[at - 1]?.to ?? -1) + 1, site.from));
	const after = runsText(runs.slice(site.to + 1, mine[at + 1]?.from ?? runs.length));
	const checkbox = site.formField?.kind === 'checkbox' || site.control?.kind === 'checkbox';
	return (checkbox ? after || before : before) || undefined;
}

/** In the body, the nearest paragraph before the field's that has text and no field. */
function previousParagraph(site: FieldSite): string | undefined {
	for (let index = site.block - 1; index >= 0; index--) {
		const block = site.blocks[index]!;
		if (block.type === 'table') return undefined;
		const text = holdsField(block) ? '' : paragraphText(block);
		if (text) return text;
	}
	return undefined;
}

/** In a cell, the nearest paragraph above the field's in the same cell that is a label. */
function previousInCell(site: FieldSite): string | undefined {
	const cell = siteCell(site);
	for (let index = (site.place?.paragraph ?? 0) - 1; index >= 0; index--) {
		const paragraph = cell?.paragraphs[index];
		const text = paragraph && !holdsField(paragraph) ? paragraphText(paragraph) : '';
		if (text) return text;
	}
	return undefined;
}

/** The label parts of one field, most general first; empty when the layout gives none. */
export function labelParts(site: FieldSite, siblings: readonly FieldSite[]): string[] {
	const parts: string[] = [];
	if (site.place) {
		parts.push(...rowLabels(site));
		if (!parts.length) {
			const header = columnHeader(site) ?? previousInCell(site);
			if (header) parts.push(header);
		}
	}
	const beside = besideText(site, siblings);
	if (beside && beside !== parts.at(-1)) parts.push(beside);
	if (!parts.length && !site.place) {
		const previous = previousParagraph(site);
		if (previous) parts.push(previous);
	}
	return parts;
}

/** The labels joined, with an ordinal added to every label that more than one field has. */
export function numberedLabels(labels: readonly string[]): string[] {
	const counts = new Map<string, number>();
	for (const label of labels) counts.set(label, (counts.get(label) ?? 0) + 1);
	const seen = new Map<string, number>();
	return labels.map((label) => {
		if (counts.get(label)! < 2) return label;
		const ordinal = (seen.get(label) ?? 0) + 1;
		seen.set(label, ordinal);
		return `${label} ${ordinal}`;
	});
}
