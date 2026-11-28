import type { LayoutFragment } from './result';
import { ptToPx } from './units';

/**
 * Word breaks a line without the width of the tags that follow an untagged content control, so the
 * line can be wider than the text area. It then draws the line with every glyph pulled in (the
 * glyph advances shrink by 0.03 to 0.06 em, a space by 0.25 pt more) so that the line ends about
 * 0.7 pt past the margin, tags at their normal width (Word 16 sweeps, issue #415).
 */
const OVERSHOOT_PT = 0.7;
const SPACE_EXTRA_PT = 0.25;
/** Narrow glyphs give up about 0.85 of what the others do. */
const NARROW = /[iljtfI.,;:'!|()]/u;
const NARROW_WEIGHT = 0.85;
const weightOf = (text: string) =>
	[...text].reduce((sum, char) => sum + (NARROW.test(char) ? NARROW_WEIGHT : 1), 0);

/** Whether `fragments`, one line, run past `widthPx` by more than Word's overshoot. */
export function compressTagLine(fragments: LayoutFragment[], widthPx: number): LayoutFragment[] {
	const last = fragments.at(-1);
	if (!last) return fragments;
	const excess = last.xPx + last.widthPx - widthPx - ptToPx(OVERSHOOT_PT);
	if (excess <= 0) return fragments;
	const glyphs = (fragment: LayoutFragment) =>
		fragment.contentControlTag || fragment.object || fragment.leader
			? 0
			: [...fragment.text].length;
	const weight = (fragment: LayoutFragment) => (glyphs(fragment) ? weightOf(fragment.text) : 0);
	const total = fragments.reduce((sum, fragment) => sum + weight(fragment), 0);
	if (!total) return fragments;
	const spaces = fragments.filter((fragment) => fragment.text === ' ').length;
	const spaceExtra = Math.min(ptToPx(SPACE_EXTRA_PT), excess / Math.max(1, spaces));
	const perGlyph = Math.max(0, (excess - spaceExtra * spaces) / total);
	let shift = 0;
	return fragments.map((fragment) => {
		const count = glyphs(fragment);
		const placed = { ...fragment, xPx: fragment.xPx - shift };
		if (!count) return placed;
		const space = fragment.text === ' ';
		const cut = perGlyph * weight(fragment) + (space ? spaceExtra : 0);
		shift += cut;
		return {
			...placed,
			widthPx: fragment.widthPx - cut,
			...(!space && { characterSpacingPx: (fragment.characterSpacingPx ?? 0) - cut / count }),
		};
	});
}
