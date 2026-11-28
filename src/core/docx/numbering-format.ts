// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { DocumentModel, Paragraph, ParagraphStyleCatalog } from './model';
import { documentParagraphs } from './cell-blocks';
import { isEmptyParagraph } from './paragraph-content';
import type { NumberingCatalog, ParagraphListLabel } from './numbering-model';
import { definedProps } from './defined-props';
import { resolveNumberingLevel } from './numbering-parse';
import { NumberingCounters, substituteLvlText } from './numbering-counters';
import { ListNumberFields } from './list-number-fields';
import { resolveStyleNumbering } from './paragraph-styles';
import {
	cardinalWords,
	letterLabel,
	ordinalNumeral,
	ordinalWords,
	romanNumeral,
} from './numbering-text';

/** Formats one placeholder value using a level's numbering format. */
export function formatListNumber(numFmt: string, value: number): string {
	switch (numFmt) {
		case 'decimal':
			return String(value);
		case 'decimalZero':
			return value < 10 && value >= 0 ? `0${value}` : String(value);
		case 'hex':
			// Native Word emits uppercase hexadecimal only within the unsigned 16-bit range.
			return Number.isInteger(value) && value >= 0 && value <= 65535
				? value.toString(16).toUpperCase()
				: '';
		case 'upperRoman':
			return romanNumeral(value);
		case 'lowerRoman':
			return romanNumeral(value).toLowerCase();
		case 'upperLetter':
			return letterLabel(value).toUpperCase();
		case 'lowerLetter':
			return letterLabel(value);
		case 'ordinal':
			return ordinalNumeral(value);
		case 'ordinalText':
			return ordinalWords(value);
		case 'cardinalText':
			return cardinalWords(value);
		case 'none':
			return '';
		case 'bullet':
			return '';
		default:
			// Unsupported/legacy formats (e.g. chicago, chineseCounting) fall back to Decimal Number.
			return String(value);
	}
}

/** Resolves the paragraph's effective `{ numId, level }`, direct or inherited through its style. */
export function resolveParagraphNumbering(
	paragraph: Paragraph,
	styleCatalog: ParagraphStyleCatalog | undefined,
	numberingCatalog?: NumberingCatalog,
): { numId: number; level: number } | undefined {
	if (paragraph.numbering)
		// `numId="0"` is Word's explicit "no numbering" override; it blocks style inheritance
		// without itself producing a marker.
		return paragraph.numbering.numId > 0 ? paragraph.numbering : undefined;
	if (!styleCatalog) return undefined;
	const inherited = resolveStyleNumbering(paragraph.style, styleCatalog);
	if (!inherited || inherited.numId <= 0) return undefined;
	if (numberingCatalog && paragraph.style && styleCatalog.styles[paragraph.style]) {
		for (let level = 0; level < 9; level++) {
			const definition = resolveNumberingLevel(numberingCatalog, String(inherited.numId), level);
			if (definition?.paragraphStyleId === paragraph.style)
				return { numId: inherited.numId, level };
		}
	}
	return inherited;
}

/** List markers by paragraph id, and the LISTNUM fields counted with the same lists. */
export interface ListNumbering {
	labels: Map<string, ParagraphListLabel>;
	listNumbers: ListNumberFields;
}

/**
 * Computes list markers for every numbered/bulleted paragraph in document order, tracking
 * counter state across the document per `numId`/level, and the LISTNUM fields that count with
 * those lists (a LISTNUM moves its list on, so later markers count it). Purely derived from the
 * model; never mutates it and is never itself part of the saved document.
 */
export function computeListNumbering(model: DocumentModel): ListNumbering {
	const catalog = model.numberingCatalog;
	const labels = new Map<string, ParagraphListLabel>();
	const counters = new NumberingCounters();
	const listNumbers = new ListNumberFields(counters, catalog);
	// The last section describes the body itself. Its last paragraph is an ordinary list item.
	const sectionBreaks = new Set(
		model.sections?.slice(0, -1).map((section) => section.endsAtBlockId),
	);
	for (const paragraph of documentParagraphs(model.blocks)) {
		if (sectionBreaks.has(paragraph.id) && isEmptyParagraph(paragraph)) continue;
		const label = catalog && paragraphLabel(paragraph, model, catalog, counters);
		if (label) labels.set(paragraph.id, label);
		listNumbers.paragraph(
			paragraph,
			label ? { numId: label.numId, level: label.level } : undefined,
		);
	}
	return { labels, listNumbers };
}

/** List markers by paragraph id (see `computeListNumbering`). */
export function computeListLabels(model: DocumentModel): Map<string, ParagraphListLabel> {
	return computeListNumbering(model).labels;
}

/** A paragraph's list marker, its list's counter moved on; undefined when it has none. */
function paragraphLabel(
	paragraph: Paragraph,
	model: DocumentModel,
	catalog: NumberingCatalog,
	counters: NumberingCounters,
): ParagraphListLabel | undefined {
	const numbering = resolveParagraphNumbering(paragraph, model.paragraphStyles, catalog);
	if (!numbering) return undefined;
	const numId = String(numbering.numId);
	const level = Math.min(8, Math.max(0, numbering.level));
	const def = resolveNumberingLevel(catalog, numId, level);
	if (!def || def.numFmt === 'none') return undefined;
	counters.advance(numId, level, def, catalog);
	const text =
		def.numFmt === 'bullet'
			? def.lvlText
			: substituteLvlText(def.lvlText, numId, level, catalog, counters, Boolean(def.isLgl));
	const picture =
		def.pictureBulletId !== undefined ? catalog.pictureBullets?.[def.pictureBulletId] : undefined;
	return {
		numId,
		level,
		text,
		suffix: def.suffix ?? 'tab',
		...definedProps({
			alignment: def.lvlJc,
			indentLeftTwips: def.indentLeftTwips,
			hangingTwips: def.hangingTwips,
			firstLineTwips: def.firstLineTwips,
			markerFormat: def.markerFormat,
			tabStopTwips: def.tabStopTwips,
			legacy: def.legacy?.enabled ? def.legacy : undefined,
			picture,
		}),
	};
}

export { displayListLabel } from './symbol-bullets';
