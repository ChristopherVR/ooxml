import { fieldResultSpans, replaceFieldResult } from '../field-result-spans';
import { fieldName } from '../field-runs';
import type { Paragraph, TextRun } from '../model';
import { quoteReference, referenceInstruction } from './instructions';
import type { Bibliography, BibliographySource, BibliographyStyle } from './source-model';

export interface CitationOptions {
	pages: string;
	suppressAuthor: boolean;
	suppressTitle: boolean;
	suppressYear: boolean;
}

export function citationCode(tag: string, lcid: number, options: CitationOptions): string {
	return [
		'CITATION',
		/^[^\s"\\]+$/u.test(tag) ? tag : quoteReference(tag),
		`\\l ${lcid}`,
		options.pages ? `\\p ${quoteReference(options.pages)}` : '',
		options.suppressAuthor ? '\\n' : '',
		options.suppressTitle ? '\\t' : '',
		options.suppressYear ? '\\y' : '',
	]
		.filter(Boolean)
		.join(' ');
}

export function readCitation(instruction: string) {
	const code = referenceInstruction(instruction);
	const tag = code?.args[0];
	if (code?.name !== 'CITATION' || !tag) return undefined;
	return {
		tag,
		pages: code.value('p') ?? '',
		suppressAuthor: code.has('n'),
		suppressTitle: code.has('t'),
		suppressYear: code.has('y'),
	};
}

export function citationSpans(runs: readonly TextRun[]) {
	return fieldResultSpans(runs).filter((field) => readCitation(field.instr));
}

/** Reuse Word's field tokenizer so an edit can keep switches such as prefix, suffix and extra tags. */
export function editCitationCode(
	instruction: string,
	options: CitationOptions,
): string | undefined {
	const code = referenceInstruction(instruction);
	const tag = code?.args[0];
	if (!code || code.name !== 'CITATION' || !tag) return undefined;
	return [
		citationCode(tag, Number(code.value('l')) || 1033, options),
		...['f', 's', 'v', 'm', '*'].flatMap((key) =>
			code.values(key).map((value) => `\\${key} ${quoteReference(value)}`),
		),
	].join(' ');
}

const surnames = (source: BibliographySource): string[] =>
	source.author?.kind === 'corporate'
		? [source.author.name]
		: (source.author?.people ?? []).map((name) => name.last || name.first).filter(Boolean);

function citationAuthor(source: BibliographySource, style: BibliographyStyle): string {
	const names = surnames(source);
	const [first = '', second = ''] = names;
	if (names.length > 2) return `${first} et al.`;
	if (names.length === 2) return `${first}${style === 'APA' ? ' & ' : ' and '}${second}`;
	return first;
}

export function citationEvaluator(
	bibliography: Bibliography,
): (instruction: string) => string | undefined {
	const sources = new Map(bibliography.sources.map((source) => [source.tag, source]));
	const style = bibliography.style ?? 'APA';
	return (instruction) => {
		if (fieldName(instruction) !== 'CITATION') return undefined;
		const code = referenceInstruction(instruction);
		const citation = readCitation(instruction);
		if (!code || !citation) return undefined;
		const tags = [citation.tag, ...code.values('m')];
		if (tags.some((tag) => !sources.has(tag))) return undefined;
		const parts = tags.map((tag) => {
			const source = sources.get(tag);
			if (!source || source.placeholder) return tag;
			const author = !citation.suppressAuthor && citationAuthor(source, style);
			const title = !citation.suppressTitle && source.fields.Title;
			const year = !citation.suppressYear && (source.fields.Year || 'n.d.');
			const pages = citation.pages;
			const subject = author || title;
			if (style === 'MLA') return [subject, pages].filter(Boolean).join(' ');
			if (style === 'Chicago')
				return [[subject, year].filter(Boolean).join(' '), pages].filter(Boolean).join(', ');
			const locator = pages && `${/[-–,]/.test(pages) ? 'pp.' : 'p.'} ${pages}`;
			return [subject, year, locator].filter(Boolean).join(', ');
		});
		return `(${code.value('f') ?? ''}${parts.join('; ')}${code.value('s') ?? ''})`;
	};
}

export function refreshCitations(
	paragraph: Paragraph,
	evaluate: ReturnType<typeof citationEvaluator>,
): Paragraph {
	let runs = paragraph.runs;
	for (const field of fieldResultSpans(runs).sort((a, b) => b.from - a.from)) {
		if (field.locked || field.removed) continue;
		const text = evaluate(field.instr);
		if (text !== undefined) runs = replaceFieldResult(runs, field, text, false) ?? runs;
	}
	return runs === paragraph.runs ? paragraph : { ...paragraph, runs };
}

function authorName(source: BibliographySource, style: BibliographyStyle): string {
	const { author } = source;
	if (!author) return '';
	if (author.kind === 'corporate') return author.name;
	const names = author.people
		.map(({ last, first, middle }, index) => {
			const given = [first, middle].filter(Boolean);
			if (style === 'APA')
				return [
					last,
					given
						.flatMap((name) => name.split(/\s+/u).map((part) => `${Array.from(part)[0]}.`))
						.join(' '),
				]
					.filter(Boolean)
					.join(', ');
			return index === 0
				? [last, given.join(' ')].filter(Boolean).join(', ')
				: [...given, last].filter(Boolean).join(' ');
		})
		.filter(Boolean);
	if (style === 'MLA' && names.length > 2) return `${names[0]}, et al.`;
	if (names.length < 2) return names[0] ?? '';
	const join = style === 'APA' ? ', & ' : names.length === 2 ? ' and ' : ', and ';
	return names.slice(0, -1).join(', ') + join + names.at(-1);
}

/** The supported common source types use the published APA, MLA and Chicago author-date patterns. */
export function bibliographyEntry(source: BibliographySource, style: BibliographyStyle): TextRun[] {
	const { fields } = source;
	const title = fields.Title || source.tag;
	const author = authorName(source, style);
	const year = fields.Year || 'n.d.';
	const container = fields.JournalName || fields.BookTitle || fields.InternetSiteTitle;
	const number = [fields.Volume, fields.Issue ? `(${fields.Issue})` : ''].filter(Boolean).join('');
	const runs: TextRun[] = [];
	const text = (value: string | undefined) => {
		if (value) runs.push({ text: value });
	};
	const period = (value: string | undefined) => {
		if (value) text(/[.?!]$/.test(value) ? `${value} ` : `${value}. `);
	};
	period(author);
	if (style === 'APA') period(`(${year})`);
	else if (style === 'Chicago') period(year);
	if (container) {
		period(style === 'APA' ? title : `“${title}.”`);
		runs.push({ text: container, italic: true });
		if (style === 'MLA')
			text(
				[
					fields.Volume && `vol. ${fields.Volume}`,
					fields.Issue && `no. ${fields.Issue}`,
					fields.Year,
					fields.Pages && `pp. ${fields.Pages}`,
				]
					.filter(Boolean)
					.map((part) => `, ${part}`)
					.join(''),
			);
		else {
			if (number) text(`${style === 'APA' ? ', ' : ' '}${number}`);
			if (fields.Pages) text(`${style === 'APA' ? ', ' : ': '}${fields.Pages}`);
		}
		text('. ');
	} else {
		runs.push({ text: title, italic: true });
		if (fields.Edition && style === 'APA') text(` (${fields.Edition} ed.)`);
		text(style === 'MLA' ? ', ' : '. ');
	}
	if (style === 'MLA' && !container)
		period([fields.Publisher, fields.Year].filter(Boolean).join(', '));
	else period(fields.Publisher);
	text(fields.URL);
	const last = runs.at(-1);
	if (last) last.text = last.text.trimEnd();
	return runs;
}
