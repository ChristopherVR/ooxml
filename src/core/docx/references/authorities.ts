import type { Paragraph, TextRun } from '../model';
import type { Twips } from '../units';
import {
	addPage,
	entryPage,
	pageRuns,
	type referenceEntries,
	type ReferenceEntry,
	type ReferencePage,
} from './entries';
import { referenceCollation, type referenceInstruction } from './instructions';

export const AUTHORITY_CATEGORIES = [
	'Cases',
	'Statutes',
	'Other Authorities',
	'Rules',
	'Treatises',
	'Regulations',
	'Constitutional Provisions',
] as const;

export type AuthorityCategory = (typeof AUTHORITY_CATEGORIES)[number];

export function authorityCode(category: number, passim: boolean, keepFormatting: boolean): string {
	return `TOA \\h \\c "${category}"${passim ? ' \\p' : ''}${keepFormatting ? '' : ' \\f'}`;
}

export function authorityParagraphs(
	collected: ReturnType<typeof referenceEntries>,
	code: NonNullable<ReturnType<typeof referenceInstruction>>,
	page: ReferencePage,
	width: Twips,
	locale?: string,
	translateCategory?: (name: AuthorityCategory) => string,
): Paragraph[] {
	const category = code.value('c') ?? '1';
	const matching = collected.entries.filter(
		({ code: entry }) => entry.name === 'TA' && (entry.value('c') ?? '1') === category,
	);
	const longByShort = new Map<string, string>();
	const groups = new Map<
		string,
		{ pages: Map<string, TextRun>; format: ReferenceEntry['format'] }
	>();
	for (const entry of matching) {
		const long = entry.code.value('l') ?? entry.code.args[0];
		const short = entry.code.value('s');
		if (!long) continue;
		if (short && !longByShort.has(short)) longByShort.set(short, long);
		if (!groups.has(long)) groups.set(long, { pages: new Map(), format: entry.format });
	}
	for (const entry of matching) {
		const group = groups.get(
			entry.code.value('l') ??
				entry.code.args[0] ??
				longByShort.get(entry.code.value('s') ?? '') ??
				'',
		);
		if (group)
			addPage(
				group.pages,
				entry.code,
				entryPage(entry, collected.bookmarks, page, code.value('g') ?? '–'),
			);
	}
	const separator = code.value('e') ?? '\t';
	const { collator } = referenceCollation(code.value('z'), locale);
	const paragraphs = [...groups]
		.sort(([a], [b]) => collator.compare(a, b))
		.map<Paragraph>(([text, { pages, format }]) => ({
			type: 'paragraph',
			id: '',
			runs: [
				{ text, ...(!code.has('f') ? format : {}) },
				...(pages.size
					? [
							{ text: separator },
							...(code.has('p') && pages.size >= 5
								? [{ text: 'passim' }]
								: pageRuns(pages, code.value('l') ?? ', ')),
						]
					: []),
			],
			...(separator === '\t'
				? { tabStops: [{ posTwips: width, align: 'right', leader: 'dot' }] }
				: {}),
		}));
	if (paragraphs.length && code.has('h')) {
		const name = AUTHORITY_CATEGORIES[Number(category) - 1];
		const heading: Paragraph = {
			type: 'paragraph',
			id: '',
			runs: [{ text: name ? (translateCategory?.(name) ?? name) : category, bold: true }],
		};
		return [heading, ...paragraphs];
	}
	return paragraphs;
}
