/**
 * Complex fields as model runs (`fieldChar`, `fieldCode` and result runs), with the results a
 * sample stores for them. Word shows a stored result until the field is updated, so every result
 * here is the value Word itself computes for the finished file (see each sample's `pages.ts`).
 */
import type { Paragraph, TextRun } from '../../../../src/core/docx/index';
import { toRuns, type Doc, type Inline, type Runs } from './builder';

/** `begin`, instruction, `separate`, the result runs and `end`; `props` goes on every run. */
export function field(instr: string, result: Inline, props: Partial<TextRun> = {}): Runs {
	return [...fieldOpen(instr, props), ...fieldResult(instr, result, props), fieldClose(props)];
}

/** The runs that open a field whose result spans several paragraphs (a table of contents). */
export function fieldOpen(instr: string, props: Partial<TextRun> = {}): Runs {
	return [
		{ ...props, text: '', fieldChar: 'begin' },
		{ ...props, text: '', fieldCode: ` ${instr} ` },
		{ ...props, text: '', fieldChar: 'separate' },
	];
}

export function fieldResult(instr: string, result: Inline, props: Partial<TextRun> = {}): Runs {
	return toRuns(result).map((run) => ({ ...props, ...run, field: { instr } }));
}

export const fieldClose = (props: Partial<TextRun> = {}): TextRun => ({
	...props,
	text: '',
	fieldChar: 'end',
});

/** A cross-reference: `REF name \h` showing `text`, the bookmarked text Word would insert. */
export const ref = (name: string, text: string, switches = '\\h', style?: string): Runs =>
	// Word copies the character style of the bookmarked text into the result: `style` names it.
	field(`REF ${name} ${switches}`, style ? { text, style } : text);

/** `PAGEREF name \h` showing the page Word puts the bookmark on. */
export const pageRef = (name: string, page: string | number, props: Partial<TextRun> = {}): Runs =>
	field(`PAGEREF ${name} \\h`, String(page), props);

/**
 * A caption paragraph: "Figure 3: text", where the number is a SEQ field. The label and number sit
 * in the `_Ref` bookmark cross-references read, and the whole paragraph in the `_Toc` bookmark the
 * list of figures (or tables) points at.
 */
export function caption(
	doc: Doc,
	label: 'Figure' | 'Table',
	number: number,
	text: Inline,
	marks: { ref: string; toc: string },
	props: Partial<Paragraph> = {},
): Paragraph {
	return doc.p(
		'Caption',
		[
			...doc.bookmark(marks.ref, [
				{ text: `${label} `, style: 'CaptionLabel' },
				// The style is on every run of the field: Word formats an updated result like the code.
				...field(`SEQ ${label} \\* ARABIC`, String(number), { style: 'CaptionLabel' }),
			]),
			{ text: '  ' },
			...toRuns(text),
		],
		{ bookmarks: [marks.toc], ...props },
	);
}
