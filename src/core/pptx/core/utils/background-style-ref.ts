/**
 * Read and write the `<p:bg><p:bgRef/></p:bg>` form of a theme background style.
 *
 * @module utils/background-style-ref
 */
import type { PptxSlideBackgroundStyleRef, XmlObject } from '../types';
import { backgroundStylePreset } from '../types/background-style';

/**
 * The preset reference a slide's own `p:bgRef` describes, or `undefined` when it is not one of
 * PowerPoint's twelve (another idx, a transformed colour, an explicit colour).
 */
export function slideBackgroundStyleRefFromNode(
	bgRef: XmlObject | undefined,
	resolvedColor: string | undefined,
	resolvedGradient: string | undefined,
): PptxSlideBackgroundStyleRef | undefined {
	if (!bgRef) {
		return undefined;
	}
	const idx = Number.parseInt(String(bgRef['@_idx'] ?? ''), 10);
	if (idx !== 1001 && idx !== 1002 && idx !== 1003) {
		return undefined;
	}
	const colour = bgRef['a:schemeClr'] as XmlObject | undefined;
	const scheme = String(colour?.['@_val'] ?? '');
	const transformed = colour && Object.keys(colour).some((key) => !key.startsWith('@_'));
	if ((scheme !== 'bg1' && scheme !== 'bg2') || transformed) {
		return undefined;
	}
	return {
		kind: 'preset',
		idx,
		scheme,
		...(resolvedColor !== undefined ? { resolvedColor } : {}),
		...(resolvedGradient ? { resolvedGradient } : {}),
	};
}

/** The `p:bg` node for a preset reference. */
export function backgroundRefNode(idx: number, scheme: 'bg1' | 'bg2'): XmlObject {
	return { 'p:bgRef': { '@_idx': String(idx), 'a:schemeClr': { '@_val': scheme } } };
}

/** The `p:bgRef` of preset `index` (1 to 12), or `undefined`. */
export function backgroundRefNodeForPreset(index: number): XmlObject | undefined {
	const preset = backgroundStylePreset(index);
	return preset ? backgroundRefNode(preset.idx, preset.scheme) : undefined;
}

function sameColour(a: string | undefined, b: string | undefined): boolean {
	const normalise = (value: string | undefined) =>
		(value ?? '').trim().replace(/^#/u, '').toUpperCase();
	return normalise(a) === normalise(b);
}

/**
 * True when the slide's painted background still equals what the reference resolved to, so the
 * reference is still what the slide shows. False once an edit (Format Background) moved it.
 */
export function backgroundStyleRefIsCurrent(
	ref: PptxSlideBackgroundStyleRef,
	slide: {
		backgroundColor?: string | undefined;
		backgroundGradient?: string | undefined;
		backgroundImage?: string | undefined;
	},
): boolean {
	return (
		sameColour(ref.resolvedColor, slide.backgroundColor) &&
		(ref.resolvedGradient ?? '') === (slide.backgroundGradient ?? '') &&
		!slide.backgroundImage
	);
}
