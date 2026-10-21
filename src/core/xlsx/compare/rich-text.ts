// Rich-text runs of a string cell. A run can be seen three ways: as written, against its cell's
// font (what it changes), and as the grid paints it (`mergeFont`, the cell font with the run's own
// settings on top). Each view hides one kind of representation-only difference: a new cell font
// under unchanged runs, a run that repeats the cell font, runs rewritten to paint the same. A run
// change is reported only where one stretch of text differs in all three.
import { mergeFont } from '../layout/style-view';
import type { Cell, Font, RichTextRun } from '../model';
import { stableKey } from '../stable-key';
import { canonicalFont, withCanonicalColors } from './format-key';

type Look = (cell: Font, run: Font | undefined) => string;

/** What a run changes in its cell's font; `{}` for a run that paints like the cell. */
const relative: Look = (cell, run) => {
	const base = canonicalFont(cell);
	const painted = canonicalFont(mergeFont(cell, run));
	const delta: Record<string, unknown> = {};
	for (const key of new Set([...Object.keys(base), ...Object.keys(painted)]) as Set<keyof Font>)
		if (stableKey(painted[key]) !== stableKey(base[key])) delta[key] = painted[key] ?? null;
	return stableKey(delta);
};

const LOOKS: Look[] = [
	// As written, `false` flags kept: in a run they switch a setting of the cell font off.
	(_, run) => stableKey(withCanonicalColors({ ...run, family: undefined })),
	relative,
	(cell, run) => stableKey(canonicalFont(mergeFont(cell, run))),
];

/** The runs of a cell's text, plain text being one run without a font. */
function runsOf(cell: Cell | undefined): readonly RichTextRun[] {
	const text = typeof cell?.value === 'string' ? cell.value : '';
	return cell?.richText?.length ? cell.richText : [{ text }];
}

/** Whether a stretch of text in run `b` (cell font `bFont`) and run `a` differs in every view. */
const differsEverywhere = (bFont: Font, b: Font | undefined, aFont: Font, a: Font | undefined) =>
	LOOKS.every((look) => look(bFont, b) !== look(aFont, a));

/** Same text: walks both run lists by character offset and compares each shared stretch. */
function spanChanged(
	b: readonly RichTextRun[],
	bFont: Font,
	a: readonly RichTextRun[],
	aFont: Font,
) {
	let i = 0;
	let j = 0;
	let bEnd = b[0]?.text.length ?? 0;
	let aEnd = a[0]?.text.length ?? 0;
	let at = 0;
	while (i < b.length && j < a.length) {
		const end = Math.min(bEnd, aEnd);
		if (end > at && differsEverywhere(bFont, b[i]?.font, aFont, a[j]?.font)) return true;
		at = end;
		if (bEnd === end) bEnd += b[++i]?.text.length ?? 0;
		if (aEnd === end) aEnd += a[++j]?.text.length ?? 0;
	}
	return false;
}

/**
 * Edited text: the sequences of run looks (neighbouring runs that look the same joined), compared
 * in every view. Text that paints like its cell has no runs in any view.
 */
function sequenceChanged(
	b: readonly RichTextRun[],
	bFont: Font,
	a: readonly RichTextRun[],
	aFont: Font,
) {
	const plain = (runs: readonly RichTextRun[], font: Font) =>
		runs.every((run) => !run.text || relative(font, run.font) === '{}');
	if (plain(b, bFont) && plain(a, aFont)) return false;
	const sequence = (runs: readonly RichTextRun[], font: Font, look: Look) => {
		const keys: string[] = [];
		for (const run of runs) {
			const key = run.text ? look(font, run.font) : undefined;
			if (key !== undefined && keys.at(-1) !== key) keys.push(key);
		}
		return keys.join('\n');
	};
	return LOOKS.every((look) => sequence(b, bFont, look) !== sequence(a, aFont, look));
}

/**
 * The runs of a cell's string are formatted differently: on some stretch of the same text, or in
 * the sequence of run looks when the text was edited too. Fonts are the two cells' own.
 */
export function richTextChanged(
	b: Cell | undefined,
	a: Cell | undefined,
	bFont: Font,
	aFont: Font,
): boolean {
	if (!b?.richText?.length && !a?.richText?.length) return false;
	const before = runsOf(b);
	const after = runsOf(a);
	// Text typed into an empty cell, or cleared, is a value change; its runs come with it.
	const textOf = (runs: readonly RichTextRun[]) => runs.map((run) => run.text).join('');
	if (!textOf(before) || !textOf(after)) return false;
	const sameText =
		typeof b?.value === 'string' &&
		b.value === a?.value &&
		textOf(before) === b.value &&
		textOf(after) === b.value;
	return sameText
		? spanChanged(before, bFont, after, aFont)
		: sequenceChanged(before, bFont, after, aFont);
}
