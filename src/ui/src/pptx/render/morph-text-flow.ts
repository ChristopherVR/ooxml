/**
 * Where a slide's PARAGRAPHS go across a morph, and the two matching rules
 * that follow from it.
 *
 * Measured on PowerPoint 16's own render (`CreateVideo`) of issue #141's deck,
 * slides 17 to 18, and on decks built to isolate the rule:
 *
 *   - A text box whose paragraphs are redistributed over SEVERAL boxes on the
 *     other slide is not carried across. The deck's left column keeps its
 *     heading and first paragraph, but one of its paragraphs moves into the
 *     right column's box; both columns fade where they stand, although one of
 *     them keeps its `a16:creationId` and the other its `p:cNvPr` id and name.
 *     Pasting the same boxes so that every paragraph stays in one box makes
 *     PowerPoint glide them, and flattening the formatting changes nothing, so
 *     the split is what decides.
 *   - A text box that says exactly the same thing on both slides is carried
 *     across even when every id and its name differ: the deck's third column
 *     ("Object 3", parked off the left edge) glides in as "Object 6".
 *
 * Only paragraphs that are DISTINCTIVE on their slide count as evidence: a
 * heading repeated in two boxes on one slide ("Lor ipsumi" heads both columns)
 * says nothing about which box it belongs to.
 *
 * @module render/morph-text-flow
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { MorphPair } from './morph-types';

/** An element's own paragraphs, whitespace-normalised, without the empty ones. */
function paragraphsOf(element: PptxElement): string[] {
	const text = (element as { text?: string }).text ?? '';
	return text
		.split(/[\n\r\v]/u)
		.map((paragraph) => paragraph.replace(/\s+/gu, ' ').trim())
		.filter((paragraph) => paragraph !== '');
}

/** Paragraph -> the one element on its slide that holds it; repeated paragraphs are left out. */
function distinctiveParagraphs(elements: readonly PptxElement[]): Map<string, PptxElement> {
	const owner = new Map<string, PptxElement | null>();
	for (const element of elements) {
		for (const paragraph of new Set(paragraphsOf(element))) {
			owner.set(paragraph, owner.has(paragraph) ? null : element);
		}
	}
	const distinctive = new Map<string, PptxElement>();
	for (const [paragraph, element] of owner) {
		if (element) {
			distinctive.set(paragraph, element);
		}
	}
	return distinctive;
}

/** Paragraph ownership on both slides, computed once per match. */
export interface MorphTextFlow {
	readonly from: ReadonlyMap<string, PptxElement>;
	readonly to: ReadonlyMap<string, PptxElement>;
}

export function morphTextFlow(
	fromElements: readonly PptxElement[],
	toElements: readonly PptxElement[],
): MorphTextFlow {
	return { from: distinctiveParagraphs(fromElements), to: distinctiveParagraphs(toElements) };
}

/**
 * Whether pairing `fromEl` with `toEl` would contradict where their paragraphs
 * go: one of `fromEl`'s distinctive paragraphs lands in a DIFFERENT box on the
 * other slide, or one of `toEl`'s comes from a different box. PowerPoint does
 * not carry such a box across, whatever identity the two share.
 */
export function splitsTextFlow(
	flow: MorphTextFlow,
	fromEl: PptxElement,
	toEl: PptxElement,
): boolean {
	const elsewhere = (
		element: PptxElement,
		own: ReadonlyMap<string, PptxElement>,
		other: ReadonlyMap<string, PptxElement>,
		partner: PptxElement,
	): boolean =>
		paragraphsOf(element).some((paragraph) => {
			if (own.get(paragraph) !== element) {
				return false;
			}
			const holder = other.get(paragraph);
			return holder !== undefined && holder !== partner;
		});
	return elsewhere(fromEl, flow.from, flow.to, toEl) || elsewhere(toEl, flow.to, flow.from, fromEl);
}

/** Fewest non-space characters a text must have to identify its box on its own. */
const MIN_TWIN_TEXT = 4;

/**
 * Pass: pair text boxes that say EXACTLY the same thing, when that text is
 * unique on both slides, however far apart they sit and whatever their ids
 * and names. A text box declares no paint and usually no shared name, so no
 * other identity pass can take it, and proximity stops at 300px.
 */
export function matchTextTwins(
	fromElements: readonly PptxElement[],
	toElements: readonly PptxElement[],
	usedFrom: Set<string>,
	usedTo: Set<string>,
): MorphPair[] {
	const byText = (elements: readonly PptxElement[]): Map<string, PptxElement[]> => {
		const map = new Map<string, PptxElement[]>();
		for (const element of elements) {
			if (element.type !== 'text') {
				continue;
			}
			const text = paragraphsOf(element).join('\n');
			if (text.replace(/\s/gu, '').length < MIN_TWIN_TEXT) {
				continue;
			}
			map.set(text, [...(map.get(text) ?? []), element]);
		}
		return map;
	};
	const toByText = byText(toElements);
	const pairs: MorphPair[] = [];
	for (const [text, froms] of byText(fromElements)) {
		const tos = toByText.get(text);
		const fromEl = froms[0];
		const toEl = tos?.[0];
		if (froms.length !== 1 || tos?.length !== 1 || !fromEl || !toEl) {
			continue;
		}
		if (usedFrom.has(fromEl.id) || usedTo.has(toEl.id)) {
			continue;
		}
		pairs.push({ fromElement: fromEl, toElement: toEl });
		usedFrom.add(fromEl.id);
		usedTo.add(toEl.id);
	}
	return pairs;
}
