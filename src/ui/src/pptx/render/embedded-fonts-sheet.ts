/**
 * embedded-fonts-sheet.ts: the managed `<style>` element that registers a
 * deck's embedded fonts as `@font-face` rules, for bindings that have no
 * reactive layer of their own to hang it on (vanilla, Svelte).
 *
 * Without it a deck's embedded faces never reach the browser: a run whose
 * family is only available as an embedded font falls through to the generic
 * tail of its CSS stack, so a binding that skipped this step painted a
 * different face than the bindings that register it. The rule text itself
 * comes from {@link buildEmbeddedFontStyles}, shared with the other bindings.
 */

import type { PptxEmbeddedFont } from 'ooxml-core/pptx';

import { buildEmbeddedFontStyles, type ObjectUrlFactory } from './embedded-fonts';

/** A binding-owned handle on the managed embedded-fonts `<style>` element. */
export interface EmbeddedFontsSheet {
	/** Replace the registered fonts; an empty list removes the element. */
	set(fonts: readonly PptxEmbeddedFont[] | null | undefined): void;
	/** Remove the element and revoke every object URL minted for it. */
	dispose(): void;
}

function hasObjectUrlSupport(): boolean {
	return (
		typeof URL !== 'undefined' &&
		typeof URL.createObjectURL === 'function' &&
		typeof URL.revokeObjectURL === 'function' &&
		typeof Blob !== 'undefined'
	);
}

const mintObjectUrl: ObjectUrlFactory = (bytes, mime) => {
	if (!hasObjectUrlSupport()) {
		return null;
	}
	// A fresh, plain-ArrayBuffer-backed copy satisfies `BlobPart`.
	const copy = new Uint8Array(bytes.length);
	copy.set(bytes);
	return URL.createObjectURL(new Blob([copy], { type: mime }));
};

/**
 * Create the handle for the `<style id="styleId">` element in `doc`. The id
 * is binding-specific so two bindings on one page never share an element.
 */
export function createEmbeddedFontsSheet(doc: Document, styleId: string): EmbeddedFontsSheet {
	let liveUrls: string[] = [];
	let lastFonts: readonly PptxEmbeddedFont[] | null | undefined;

	const revoke = (urls: readonly string[]): void => {
		if (hasObjectUrlSupport()) {
			for (const url of urls) {
				URL.revokeObjectURL(url);
			}
		}
	};
	const remove = (): void => {
		doc.getElementById(styleId)?.remove();
	};

	return {
		set(fonts) {
			// The same list again changes nothing (the store re-emits on other fields).
			if (fonts === lastFonts) {
				return;
			}
			lastFonts = fonts;
			const previous = liveUrls;
			const { fontFaceCss, objectUrls } = buildEmbeddedFontStyles(fonts, mintObjectUrl);
			liveUrls = objectUrls;
			if (fontFaceCss.length === 0) {
				remove();
			} else {
				let style = doc.getElementById(styleId) as HTMLStyleElement | null;
				if (!style) {
					style = doc.createElement('style');
					style.id = styleId;
					doc.head.appendChild(style);
				}
				style.textContent = fontFaceCss;
			}
			revoke(previous);
		},
		dispose() {
			remove();
			revoke(liveUrls);
			liveUrls = [];
			lastFonts = undefined;
		},
	};
}
