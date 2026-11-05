/**
 * Whether a matched pair is a text box whose WORDS were replaced without it
 * leaving its slot - the case PowerPoint dissolves in place rather than moving.
 *
 * @module render/morph-text-slot
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import { boxOverlapRatio } from './morph-flatten';

/**
 * Fraction of the union the two boxes must share to count as the same slot.
 *
 * A paragraph PowerPoint re-fits around its own wording can move a fair way as
 * the panel reflows: on the wheel deck's slide 7 -> 8 the "Challenge" line's box
 * shifts 17.75px up a 55px-tall box and its width changes, which drops the
 * overlap to 0.487 - just under the 0.5 that `morph-flatten` uses for group
 * children. It is still the same paragraph dissolving in place, so it was
 * stretching its glyphs by the 1.56% its box width changed (issue #161's "text
 * moving") instead of dissolving. A genuine relocation (measured: a text box
 * that moved 460px) shares no box at all and scores ~0, so 0.4 keeps that
 * interpolating while catching a re-fit this size.
 */
const SAME_SLOT_OVERLAP = 0.4;

/** An element's own words, whitespace-normalised. */
function ownText(element: PptxElement): string {
	return ((element as { text?: string }).text ?? '').replace(/\s+/gu, ' ').trim();
}

/** An unrotated box entirely outside the slide's top or left edge. */
function offstageTopLeft(element: PptxElement): boolean {
	return (
		(element.rotation ?? 0) === 0 &&
		(element.x + element.width <= 0 || element.y + element.height <= 0)
	);
}

/**
 * A size-only text change that can fly as ONE copy, because one end of the
 * flight is entirely outside the top or left edge where the switch between
 * the two layouts cannot be seen. This avoids superimposing two layouts of the
 * same text during its flight. Restrict this to size-only run changes: a
 * change in wording, font, colour or emphasis still needs a cross-dissolve.
 *
 * Either end qualifies. Issue #141's slides 17 -> 18 bring a parked paragraph
 * IN from off the left edge; played backwards (18 -> 17) the same paragraph
 * flies OUT to that parking spot, and checking only the source left the
 * reverse cross-dissolving two layouts across the slide.
 *
 * @see morphOffstageTextAdoptsTarget for which layout the single copy paints.
 */
export function morphOffstageTextSizeChange(from: PptxElement, to: PptxElement): boolean {
	return (offstageTopLeft(from) || offstageTopLeft(to)) && sizeOnlyTextChange(from, to);
}

/**
 * Whether the single travelling copy of {@link morphOffstageTextSizeChange}
 * paints the DESTINATION layout: only when the source is the off-stage end,
 * so the switch happens before the text enters. When the destination is the
 * off-stage end the copy keeps its own layout until it has left the slide.
 */
export function morphOffstageTextAdoptsTarget(from: PptxElement, to: PptxElement): boolean {
	return offstageTopLeft(from) && sizeOnlyTextChange(from, to);
}

/** The same words in the same runs, differing only in some run's font size. */
function sizeOnlyTextChange(from: PptxElement, to: PptxElement): boolean {
	if (
		from.type !== 'text' ||
		to.type !== 'text' ||
		!hasTextProperties(from) ||
		!hasTextProperties(to) ||
		JSON.stringify(from.shapeStyle ?? {}) !== JSON.stringify(to.shapeStyle ?? {}) ||
		!from.text ||
		from.text !== to.text
	) {
		return false;
	}
	const a = from.textSegments ?? [];
	const b = to.textSegments ?? [];
	if (a.length === 0 || a.length !== b.length) return false;
	let sizeChanged = false;
	for (let i = 0; i < a.length; i++) {
		const old = a[i];
		const next = b[i];
		if (!old || !next || old.text !== next.text) return false;
		const x = old.style;
		const y = next.style;
		if (x?.fontSize !== y?.fontSize) sizeChanged = true;
		if (
			x?.fontFamily !== y?.fontFamily ||
			x?.color !== y?.color ||
			(x?.colorRef?.alpha ?? 1) !== (y?.colorRef?.alpha ?? 1) ||
			Boolean(x?.bold) !== Boolean(y?.bold) ||
			Boolean(x?.italic) !== Boolean(y?.italic) ||
			Boolean(x?.underline) !== Boolean(y?.underline)
		)
			return false;
	}
	return sizeChanged;
}

/**
 * Whether a matched pair of TEXT BOXES holds different wording in the same slot.
 *
 * Such a pair is animated as a pure dissolve: no translation, no scale, each
 * half painted at its own geometry with complementary opacity. Everywhere else
 * a matched pair interpolates its whole box, which is right for a shape - but a
 * text box's box is a container that PowerPoint re-fits around whatever it now
 * says, and its glyphs are laid out inside that box rather than scaled with it.
 * Interpolating it therefore stretches the wording by the amount the WORDS
 * changed length, which is never something PowerPoint shows.
 *
 * Measured on PowerPoint 16's own render (`CreateVideo`, 62.5fps):
 *
 *   - A text box whose wording changed while its box doubled in width dissolves
 *     glyph over glyph with the type at a constant size, still on its left
 *     margin: the box grew, the text did not.
 *   - The wheel deck's centre paragraphs (issue #160) re-fit by 11px and 12px
 *     between topic slides. Every frame of PowerPoint's transition is a blend of
 *     the two end states with a residual under 1.1/255, which no scaling or
 *     shifting of either half could produce.
 *   - A text box that genuinely MOVES (460px, wording changed too) travels the
 *     whole way while its glyphs cross-dissolve, so distance has to keep the
 *     interpolation. Hence the slot test rather than a blanket rule.
 *
 * @param fromElement - The outgoing half of the pair.
 * @param toElement - The incoming half.
 * @returns True when the pair should dissolve where it stands.
 */
export function morphTextReplacedInSlot(fromElement: PptxElement, toElement: PptxElement): boolean {
	if (fromElement.type !== 'text' || toElement.type !== 'text') {
		return false;
	}
	const from = ownText(fromElement);
	const to = ownText(toElement);
	if (from === '' || to === '' || from === to) {
		return false;
	}
	return boxOverlapRatio(fromElement, toElement) >= SAME_SLOT_OVERLAP;
}
