/**
 * What the shared-clone pictures (`pdf-vector-raster-batch.ts`) need from the LIVE stage beyond
 * its computed styles. A capture of the live stage hides the stage and lets the cascade decide
 * the rest, and under the cascade a `visibility` an element declares in its own `style` beats the
 * hidden stage it inherits from; computed styles alone cannot tell such a declaration from an
 * inherited value, so the declarations are carried into the clone on a temporary attribute.
 */
import { withAuthoredStrokeWidths } from '../render/device-pixel-stroke';
import { prepareExportClone } from '../render/export-clone';
import {
	buildInlineStyleText,
	type ComputedStyleReader,
} from '../render/foreign-object-style-inline';
import { isolateLive } from './pdf-vector-raster';

/** The `visibility` a live stage element declares in its own `style`, carried by its clone. */
export const RASTER_VISIBILITY_ATTRIBUTE = 'data-pdf-raster-visibility';

/**
 * Mark every element of `stage` that declares a `visibility` in its own `style` with
 * {@link RASTER_VISIBILITY_ATTRIBUTE}. Returns the function that takes the marks off again (call
 * it in a `finally` once the clone is taken).
 */
export function markDeclaredVisibility(stage: HTMLElement): () => void {
	const marked: Element[] = [];
	const unmark = () => {
		for (const el of marked) el.removeAttribute(RASTER_VISIBILITY_ATTRIBUTE);
	};
	try {
		for (const el of [stage, ...stage.querySelectorAll('*')]) {
			const declared = (el as HTMLElement).style?.getPropertyValue('visibility');
			if (!declared) continue;
			el.setAttribute(RASTER_VISIBILITY_ATTRIBUTE, declared);
			marked.push(el);
		}
	} catch (error) {
		unmark();
		throw error;
	}
	return unmark;
}

/**
 * The declarations a `shadow` picture of `target` gives it, `!important` each and escaped for a
 * double-quoted attribute: its computed style while the live stage is isolated for that capture
 * (so values derived from its now transparent `color` follow, as in a capture), prepared the way
 * the capture clone is.
 */
export function shadowTargetStyle(
	stage: HTMLElement,
	target: Element,
	readComputedStyle: ComputedStyleReader = (el) => window.getComputedStyle(el),
): string {
	const restoreLive = isolateLive(stage, target, 'shadow');
	let text: string;
	try {
		text = withAuthoredStrokeWidths(stage, () => buildInlineStyleText(readComputedStyle(target)));
	} finally {
		restoreLive();
	}
	const scratch = target.cloneNode(false) as HTMLElement;
	scratch.setAttribute('style', text);
	prepareExportClone(scratch);
	const out: string[] = [];
	for (let i = 0; i < scratch.style.length; i++) {
		const property = scratch.style.item(i);
		out.push(`;${property}:${scratch.style.getPropertyValue(property)}!important`);
	}
	return out
		.join('')
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;');
}
