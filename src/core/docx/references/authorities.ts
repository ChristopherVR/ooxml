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
import type { referenceInstruction } from './instructions';

export function authorityCode(category: number, passim: boolean, keepFormatting: boolean): string {
	return `TOA \\c "${category}"${passim ? ' \\p' : ''}${keepFormatting ? '' : ' \\f'}`;
}

export function authorityParagraphs(
	collected: ReturnType<typeof referenceEntries>,
	code: NonNullable<ReturnType<typeof referenceInstruction>>,
	page: ReferencePage,
	width: Twips,
): Paragraph[] {
	const matching = collected.entries.filter(
		({ code: entry }) =>
			entry.name === 'TA' && (entry.value('c') ?? '1') === (code.value('c') ?? '1'),
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
	return [...groups]
		.sort(([a], [b]) => a.localeCompare(b))
		.map(([text, { pages, format }]) => ({
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
}
