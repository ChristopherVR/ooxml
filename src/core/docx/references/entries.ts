import { mapBlockParagraphs } from '../document-paragraphs';
import { fieldResultSpans } from '../field-result-spans';
import { reviewSourceLength } from '../layout/review-display';
import type { DocumentModel, TextRun } from '../model';
import { referenceInstruction } from './instructions';

export type ReferencePage = (
	paragraphId: string,
	run?: number,
	end?: boolean,
) => string | undefined;

export interface ReferenceEntry {
	paragraphId: string;
	run: number;
	code: NonNullable<ReturnType<typeof referenceInstruction>>;
	format?: Pick<TextRun, 'bold' | 'italic' | 'underline'>;
}

export function referenceEntries(model: DocumentModel) {
	const entries: ReferenceEntry[] = [];
	const bookmarks = new Map<string, { start: string; end: string }>();
	const bookmarkEnds = new Map(model.sourceBookmarkEnds);
	mapBlockParagraphs(model.blocks, (paragraph) => {
		for (const name of paragraph.bookmarks ?? [])
			if (!bookmarks.has(name))
				bookmarks.set(name, { start: paragraph.id, end: bookmarkEnds.get(name) ?? paragraph.id });
		if (entries.length >= 10_000) return paragraph;
		let offset = 0;
		const offsets = paragraph.runs.map((run) => {
			const start = offset;
			offset += reviewSourceLength(run);
			return start;
		});
		for (const span of fieldResultSpans(paragraph.runs, { resultless: true })) {
			const marker = paragraph.runs[span.from];
			if (marker?.revision?.kind === 'delete' || marker?.revision?.kind === 'moveFrom') continue;
			const code = referenceInstruction(span.instr);
			if (code?.name !== 'XE' && code?.name !== 'TA') continue;
			const run = paragraph.runs[span.code ?? span.from];
			const format: ReferenceEntry['format'] = {};
			if (run?.bold) format.bold = true;
			if (run?.italic) format.italic = true;
			if (run?.underline) format.underline = true;
			entries.push({ paragraphId: paragraph.id, run: offsets[span.from] ?? 0, code, format });
			if (entries.length >= 10_000) break;
		}
		return paragraph;
	});
	return { entries, bookmarks };
}

export function entryPage(
	entry: ReferenceEntry,
	bookmarks: ReturnType<typeof referenceEntries>['bookmarks'],
	page: ReferencePage,
	separator: string,
): string | undefined {
	const range = entry.code.value('r');
	if (!range) return page(entry.paragraphId, entry.run);
	const target = bookmarks.get(range);
	if (!target) return undefined;
	const from = page(target.start);
	const to = page(target.end, undefined, true);
	return from && to && from !== to ? from + separator + to : from;
}

export function pageReference(format: ReferenceEntry['code'], text: string): TextRun {
	return {
		text,
		...(format.has('b') ? { bold: true } : {}),
		...(format.has('i') ? { italic: true } : {}),
	};
}

export function addPage(
	runs: Map<string, TextRun>,
	code: ReferenceEntry['code'],
	page?: string,
): void {
	if (!page) return;
	const previous = runs.get(page);
	runs.set(page, { ...previous, ...pageReference(code, page) });
}

export function pageRuns(runs: Map<string, TextRun>, separator: string): TextRun[] {
	return [...runs.values()].flatMap((run, index) => (index ? [{ text: separator }, run] : [run]));
}
