import type { Paragraph, TextRun } from '../model';
import { signedTwips, twips, type Twips } from '../units';
import { addPage, entryPage, pageRuns, type referenceEntries, type ReferencePage } from './entries';
import { quoteReference, REFERENCE_LCID, type referenceInstruction } from './instructions';

export const INDEX_LEADERS = ['none', 'dot', 'hyphen', 'underscore'] as const;
export type IndexLeader = (typeof INDEX_LEADERS)[number];

export interface IndexOptions {
	columns: number;
	runIn: boolean;
	rightAlign: boolean;
	leader: IndexLeader;
	lcid: number;
}

export function indexCode(options: IndexOptions): string {
	return [
		'INDEX',
		`\\c "${Math.max(1, Math.min(4, Math.floor(options.columns) || 1))}"`,
		'\\h "A"',
		`\\z ${options.lcid}`,
		options.runIn ? '\\r' : '',
		options.rightAlign ? `\\e ${quoteReference('\t')}` : '',
	]
		.filter(Boolean)
		.join(' ');
}

type IndexBranch = {
	pages: Map<string, TextRun>;
	see: Set<string>;
	children: Map<string, IndexBranch>;
};

const branch = (): IndexBranch => ({ pages: new Map(), see: new Set(), children: new Map() });

function entryWords(text: string): string[] {
	const words: string[] = [''];
	for (let index = 0; index < text.length; index++) {
		const char = text[index];
		if (char === ':' && words.length < 9) words.push('');
		else {
			if (char === '\\' && /[:\\]/.test(text[index + 1] ?? '')) index++;
			words[words.length - 1] += text[index] ?? '';
		}
	}
	return words.map((word) => word.trim()).filter(Boolean);
}

type Code = NonNullable<ReturnType<typeof referenceInstruction>>;

export function indexParagraphs(
	collected: ReturnType<typeof referenceEntries>,
	code: Code,
	page: ReferencePage,
	width: Twips,
	leader: IndexLeader,
): Paragraph[] {
	const locale = Object.keys(REFERENCE_LCID).find(
		(lang) => REFERENCE_LCID[lang] === Number(code.value('z')),
	);
	const collator = new Intl.Collator(locale, { sensitivity: 'base' });
	const columns = Math.min(4, Math.max(1, Number(code.value('c')) || 1));
	const separator = code.value('e') ?? ', ';
	const paragraphs: Paragraph[] = [];
	const root = branch();
	for (const entry of collected.entries) {
		if (
			entry.code.name !== 'XE' ||
			(entry.code.value('f') ?? 'i').toLowerCase() !== (code.value('f') ?? 'i').toLowerCase()
		)
			continue;
		const words = entryWords(entry.code.args[0] ?? '');
		if (!words.length) continue;
		let current = root;
		for (const word of words) {
			const next = current.children.get(word) ?? branch();
			current.children.set(word, next);
			current = next;
		}
		const see = entry.code.value('t');
		if (see) current.see.add(see);
		else
			addPage(
				current.pages,
				entry.code,
				entryPage(entry, collected.bookmarks, page, code.value('g') ?? '–'),
			);
	}
	const ordered = (node: IndexBranch) =>
		[...node.children].sort(([a], [b]) => collator.compare(a, b));
	const labelRuns = (label: string, node: IndexBranch, inline: boolean): TextRun[] => [
		{ text: label },
		...(node.pages.size
			? [{ text: inline ? ' ' : separator }, ...pageRuns(node.pages, code.value('l') ?? ', ')]
			: []),
		...(node.see.size ? [{ text: (code.value('k') ?? '. ') + [...node.see].join('; ') }] : []),
	];
	const add = (runs: TextRun[], depth: number) =>
		paragraphs.push({
			type: 'paragraph',
			id: '',
			runs,
			...(depth ? { indentLeftTwips: signedTwips(depth * 240) } : {}),
			...(separator === '\t'
				? {
						tabStops: [
							{
								posTwips: twips(Math.max(1, Math.round((width - (columns - 1) * 360) / columns))),
								align: 'right',
								leader,
							},
						],
					}
				: {}),
		});
	const runIn = (node: IndexBranch): TextRun[] =>
		ordered(node).flatMap(([word, child], index) => [
			{ text: index ? '; ' : ': ' },
			...labelRuns(word, child, true),
			...runIn(child),
		]);
	const rows = (label: string, node: IndexBranch, depth: number) => {
		const runs = labelRuns(label, node, false);
		if (code.has('r')) runs.push(...runIn(node));
		add(runs, depth);
		if (!code.has('r')) for (const [word, child] of ordered(node)) rows(word, child, depth + 1);
	};
	let group: string | undefined;
	for (const [word, node] of ordered(root)) {
		const first = Array.from(word.normalize('NFC'))[0] ?? '';
		const upper = /\p{L}/u.test(first) ? first.toLocaleUpperCase(locale) : '#';
		if (code.has('h') && (group === undefined || collator.compare(upper, group) !== 0)) {
			group = upper;
			add([{ text: (code.value('h') ?? 'A').replace(/A/i, upper), bold: true }], 0);
		}
		rows(word, node, 0);
	}
	return paragraphs;
}
