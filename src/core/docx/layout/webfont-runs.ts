/**
 * Word's text reduced to the format-neutral runs the webfont fallback (`ooxml-core/text`,
 * `WebfontRun`) loads faces for. Each run carries the family it is drawn in after document
 * defaults, paragraph, table and character styles and the theme fonts, split by script the way
 * layout splits it (east Asian and complex script text take their own fonts), so only the faces
 * a document really draws are requested.
 *
 * Covers the body, tables (including their style's run formatting and cells' nested source
 * blocks), headers and footers, footnotes, endnotes and their separators, list markers, symbols
 * drawn in their own font and the plain text of text boxes (drawn in the default font). DOM-free.
 *
 * @module docx/layout/webfont-runs
 */
import { segmentByFont } from '../../text/script-fonts';
import type { WebfontRun } from '../../text/webfonts';
import { documentBlockLists } from '../document-paragraphs';
import type { Block, DocumentModel, Paragraph, Table } from '../model';
import { computeListLabels } from '../numbering-format';
import { resolveTableStyleFormatting } from '../resolve-table';
import { resolveRunFormatting } from '../run-formatting';
import type { RunFormatting } from '../run-style-model';
import { symbolOnly, symbolSegments } from '../special-characters';
import { DEFAULT_FONT_FAMILY } from './units';

type Push = (run: WebfontRun) => void;

/** Every run of `model` as a `WebfontRun` (family, emphasis and the characters drawn). */
export function collectDocxWebfontRuns(model: DocumentModel): WebfontRun[] {
	const runs: WebfontRun[] = [];
	const theme = model.theme;
	const labels = computeListLabels(model);
	const push: Push = (run) => {
		if (run.text !== '') runs.push(run);
	};
	const familyOf = (formatting: RunFormatting) => {
		const role = formatting.fontTheme?.ascii ?? formatting.fontTheme?.hAnsi;
		return formatting.fontFamily ?? (role && theme ? theme.fonts[role]?.latin : undefined);
	};
	/** The font an empty paragraph (and the caret in it) is drawn in: no text, but a face to load. */
	const emptyRun = (formatting: RunFormatting): void => {
		runs.push({
			family: familyOf(formatting) ?? DEFAULT_FONT_FAMILY,
			bold: formatting.bold === true,
			italic: formatting.italic === true,
			text: '',
		});
	};

	const emit = (text: string, formatting: RunFormatting, family: string | undefined): void => {
		const latin = family || DEFAULT_FONT_FAMILY;
		const eaRole = formatting.fontTheme?.eastAsia;
		const csRole = formatting.fontTheme?.cs;
		const bold = formatting.bold === true;
		const italic = formatting.italic === true;
		const fonts = {
			latin,
			eastAsia:
				formatting.fontFamilyEastAsia || (eaRole && theme?.fonts[eaRole]?.eastAsia) || latin,
			complexScript:
				formatting.fontFamilyComplexScript ||
				(csRole && theme?.fonts[csRole]?.complexScript) ||
				latin,
			symbol: latin,
		};
		for (const piece of segmentByFont(text, fonts))
			push({ family: piece.fontFamily, bold, italic, text: piece.text });
	};

	const paragraphRuns = (paragraph: Paragraph, tableStyleRun?: RunFormatting): void => {
		const context = {
			runCatalog: model.characterStyles,
			paragraphCatalog: model.paragraphStyles,
			paragraphStyleId: paragraph.style,
			...(tableStyleRun ? { tableStyleRun } : {}),
		};
		const label = labels.get(paragraph.id);
		const firstFormatting = resolveRunFormatting(paragraph.runs[0] ?? { text: '' }, context);
		if (label && !label.picture) {
			const marker = label.markerFormat;
			push({
				family: marker?.fontFamily ?? familyOf(firstFormatting) ?? DEFAULT_FONT_FAMILY,
				bold: marker?.bold ?? firstFormatting.bold === true,
				italic: marker?.italic ?? firstFormatting.italic === true,
				text: label.text,
			});
		}
		if (paragraph.runs.every((run) => !run.text)) emptyRun(firstFormatting);
		for (const source of paragraph.runs) {
			for (const run of symbolSegments(source)) {
				const formatting = resolveRunFormatting(run, context);
				if (formatting.vanish) continue;
				const textBox = run.image?.textBoxText;
				if (textBox) emit(textBox.join('\n'), {}, DEFAULT_FONT_FAMILY);
				if (!run.text || run.image) continue;
				// Word draws a symbol (`w:sym`) in its own font, whatever the run's font is.
				const family = (symbolOnly(run) && run.symbol?.font) || familyOf(formatting);
				emit(formatting.caps ? run.text.toUpperCase() : run.text, formatting, family);
			}
		}
	};

	const tableRuns = (table: Table): void => {
		const rowCount = table.rows.length;
		const columnCount = Math.max(
			0,
			...table.rows.map((row) => row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0)),
		);
		table.rows.forEach((row, rowIndex) => {
			let column = 0;
			for (const cell of row) {
				const tableStyleRun = table.style
					? resolveTableStyleFormatting(
							table.style,
							model.tableStyles,
							table.look,
							rowIndex,
							rowCount,
							column,
							columnCount,
						).run
					: undefined;
				for (const paragraph of cell.paragraphs) paragraphRuns(paragraph, tableStyleRun);
				const nestedBlocks = (cell.sourceBlocks ?? []).filter(
					(block): block is Table => block.type === 'table',
				);
				for (const block of nestedBlocks) tableRuns(block);
				// Without the source blocks only the nested tables' plain text is known.
				if (nestedBlocks.length === 0)
					for (const nested of cell.nestedTables ?? [])
						for (const nestedRow of nested.rows)
							for (const nestedCell of nestedRow) emit(nestedCell.text, {}, DEFAULT_FONT_FAMILY);
				column += cell.gridSpan ?? 1;
			}
		});
	};

	const blockRuns = (blocks: readonly Block[]): void => {
		for (const block of blocks) {
			if (block.type === 'paragraph') paragraphRuns(block);
			else tableRuns(block);
		}
	};

	for (const blocks of documentBlockLists(model)) blockRuns(blocks);
	// The document's default font, which new text is typed in even when no run uses it yet.
	emptyRun(
		resolveRunFormatting(
			{ text: '' },
			{ runCatalog: model.characterStyles, paragraphCatalog: model.paragraphStyles },
		),
	);
	return runs;
}
