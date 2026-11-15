// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// AUTONUM, AUTONUMLGL and AUTONUMOUT (ECMA-376 Part 1, 17.16.5.4 to 17.16.5.6): numbers counted
// by the paragraphs that hold one of these fields. Word's output, read from the documents in
// `__fixtures__/computed-fields`, shows how: a heading (a paragraph with an outline level)
// counts at its level and a body paragraph one level below the numbered heading before it, headings and
// body paragraphs keep counters of their own, a numbered heading restarts every deeper counter,
// and every such field of one paragraph shows that paragraph's number.
import { documentParagraphs } from './cell-blocks';
import { fieldResultSpans } from './field-result-spans';
import { fieldName } from './field-runs';
import { applyFieldFormat, parseFieldInstruction } from './field-format';
import type { DocumentModel, Paragraph } from './model';
import { formatListNumber } from './numbering-format';
import { resolveParagraphFormatting } from './paragraph-styles';
import { OUTLINE_NUMBERING } from './outline-numbering';

/** The field names counted together. */
export const AUTO_NUMBER_FIELDS: ReadonlySet<string> = new Set([
	'AUTONUM',
	'AUTONUMLGL',
	'AUTONUMOUT',
]);

/** What a paragraph's automatic number fields count from. */
export interface AutoNumberState {
	/** One-based level: a heading's outline level, else one below the numbered heading before it. */
	level: number;
	heading: boolean;
	/** The heading counters of levels 1 to `level - 1`, then the paragraph's own count. */
	counts: number[];
}

/** Whether a paragraph holds an automatic number field. */
function counts(paragraph: Paragraph): boolean {
	return fieldResultSpans(paragraph.runs, { resultless: true }).some((field) =>
		AUTO_NUMBER_FIELDS.has(fieldName(field.instr)),
	);
}

/** The automatic number state of every paragraph of the main document that holds such a field. */
export function autoNumberStates(model: DocumentModel): Map<string, AutoNumberState> {
	const states = new Map<string, AutoNumberState>();
	const headings = new Array<number>(11).fill(0);
	const bodies = new Array<number>(11).fill(0);
	let lastHeading = 0;
	for (const paragraph of documentParagraphs(model.blocks)) {
		const outline = model.paragraphStyles
			? resolveParagraphFormatting(paragraph, model.paragraphStyles).outlineLevel
			: paragraph.outlineLevel;
		const heading = outline !== undefined && outline >= 1 && outline <= 9;
		const level = heading ? outline : lastHeading + 1;
		// Only a heading that holds such a field sets the level of the body paragraphs after it.
		if (!counts(paragraph)) continue;
		if (heading) lastHeading = outline;
		const own = heading ? headings : bodies;
		own[level]!++;
		for (let deeper = level + 1; deeper <= 10; deeper++) {
			bodies[deeper] = 0;
			if (heading) headings[deeper] = 0;
		}
		states.set(paragraph.id, {
			level,
			heading,
			counts: [...headings.slice(1, level), own[level]!],
		});
	}
	return states;
}

/**
 * What an automatic number field shows for a paragraph's state, or undefined for another field
 * or switches this does not evaluate. AUTONUM draws the paragraph's count in its `\*` format,
 * AUTONUMLGL every level's (`\e` drops the last separator), both followed by `\s`'s first
 * character (a period by default); AUTONUMOUT draws a heading in outline numbering and a body
 * paragraph as `1.`, whatever its switches.
 */
export function autoNumberResult(instr: string, state: AutoNumberState): string | undefined {
	const field = parseFieldInstruction(instr, { s: 'arg', e: 'flag' });
	if (!field || !AUTO_NUMBER_FIELDS.has(field.name)) return undefined;
	const own = state.counts.at(-1)!;
	if (field.name === 'AUTONUMOUT') {
		const [open, numFmt, close] = state.heading
			? OUTLINE_NUMBERING[state.level - 1]!
			: ['', 'decimal', '.'];
		return `${open}${formatListNumber(numFmt, own)}${close}`;
	}
	const given = field.switches.get('s');
	const separator = typeof given === 'string' && given ? [...given][0]! : '.';
	if (field.name === 'AUTONUMLGL')
		return state.counts.join(separator) + (field.switches.has('e') ? '' : separator);
	return applyFieldFormat(String(own), field.format) + separator;
}
