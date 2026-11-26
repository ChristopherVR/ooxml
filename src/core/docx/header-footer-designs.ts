// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Insert > Header and Footer galleries: the designs this package offers, written from scratch as
// plain paragraphs with tab stops, content controls (the document title, the author and a date
// picker, showing their placeholder until filled) and the PAGE field.
import type { DocumentPropertySet } from '../opc/properties/types';
import type { ContentControl } from './content-control-model';
import { freshControlId } from './content-control-model';
import { formatWordDate } from './field-date';
import type { Block, Paragraph, TextRun } from './model';
import { signedTwips } from './units';

export const HEADER_FOOTER_DESIGNS = [
	'blank',
	'threeColumns',
	'title',
	'titleAndAuthor',
	'centeredTitle',
	'authorAndPage',
	'dateAndPage',
] as const;
export type HeaderFooterDesign = (typeof HEADER_FOOTER_DESIGNS)[number];

export const isHeaderFooterDesign = (value: string): value is HeaderFooterDesign =>
	(HEADER_FOOTER_DESIGNS as readonly string[]).includes(value);

export interface HeaderFooterDesignContext {
	newId: () => string;
	/** The text area's width in twips: the right tab stop sits on its right edge. */
	contentWidthTwips: number;
	properties: DocumentPropertySet;
	/** Today, for the date picker; defaults to now. */
	now?: Date;
	/** The placeholder texts and control titles, in the display language. */
	labels?: { title: string; author: string; date: string };
}

const TAB: TextRun = { text: '\t' };
const DATE_PICTURE = 'MMMM d, yyyy';

/** The local calendar day as an `xsd:dateTime`, so it matches what the control shows. */
const localDay = (date: Date): string => {
	const two = (n: number) => String(n).padStart(2, '0');
	return `${date.getFullYear()}-${two(date.getMonth() + 1)}-${two(date.getDate())}T00:00:00Z`;
};

/** The blocks a design puts in a header or footer (`blank` is one empty paragraph). */
export function headerFooterDesignBlocks(
	design: HeaderFooterDesign,
	context: HeaderFooterDesignContext,
): Block[] {
	const { newId, properties } = context;
	const now = context.now ?? new Date();
	const labels = context.labels ?? { title: 'Document title', author: 'Author name', date: 'Date' };
	const used = new Set<string>();
	const control = (extra: Partial<ContentControl>): ContentControl => {
		const id = freshControlId(used);
		used.add(id);
		return { id, level: 'inline', kind: 'text', ...extra };
	};
	/** A text control that shows `value`, or its bracketed prompt while there is none. */
	const prompt = (alias: string, value: string | undefined, label: string): TextRun => {
		const empty = !value?.trim();
		return {
			text: empty ? `[${label}]` : value!,
			contentControl: control({
				alias,
				tag: alias,
				...(empty ? { showingPlaceholder: true, placeholderText: `[${label}]` } : {}),
			}),
		};
	};
	const title = (): TextRun => prompt('Title', properties.core.title, labels.title);
	const author = (): TextRun => prompt('Author', properties.core.creator, labels.author);
	const date = (): TextRun => ({
		text: formatWordDate(now, DATE_PICTURE),
		contentControl: control({
			kind: 'date',
			alias: labels.date,
			tag: 'Date',
			dateFormat: DATE_PICTURE,
			dateLocale: 'en-US',
			fullDate: localDay(now),
		}),
	});
	const page: TextRun = { text: '1', field: { instr: ' PAGE ', simple: true } };
	const right = signedTwips(Math.max(1440, Math.round(context.contentWidthTwips)));
	const center = signedTwips(Math.round(right / 2));
	const threeTabs = [
		{ posTwips: center, align: 'center' as const },
		{ posTwips: right, align: 'right' as const },
	];
	const paragraph = (runs: TextRun[], extra: Partial<Paragraph> = {}): Paragraph => ({
		type: 'paragraph',
		id: newId(),
		runs,
		...extra,
	});
	switch (design) {
		case 'blank':
			return [paragraph([])];
		case 'threeColumns':
			return [paragraph([TAB, TAB], { tabStops: threeTabs })];
		case 'title':
			return [paragraph([{ ...title(), bold: true }])];
		case 'titleAndAuthor':
			return [paragraph([title(), TAB, TAB, author()], { tabStops: threeTabs })];
		case 'centeredTitle':
			return [paragraph([{ ...title(), bold: true }], { align: 'center' })];
		case 'authorAndPage':
			return [paragraph([author(), TAB, TAB, page], { tabStops: threeTabs })];
		case 'dateAndPage':
			return [paragraph([date(), TAB, TAB, page], { tabStops: threeTabs })];
	}
}
