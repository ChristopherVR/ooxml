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
	/** Host-supplied `@font-face` text appended after the deck's own rules. */
	setExtraCss(css: string): void;
	/** The deck's `@font-face` text from the last `set` (without the extra CSS). */
	readonly fontFaceCss: string;
	/** Distinct substitution-resolved family strings of the last `set`. */
	readonly fontFamilies: readonly string[];
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
	let embeddedCss = '';
	let extraCss = '';
	let families: string[] = [];

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

	const paint = (): void => {
		const css = [embeddedCss, extraCss].filter(Boolean).join('\n');
		if (css.length === 0) {
			remove();
			return;
		}
		let style = doc.getElementById(styleId) as HTMLStyleElement | null;
		if (!style) {
			style = doc.createElement('style');
			style.id = styleId;
			doc.head.appendChild(style);
		}
		style.textContent = css;
	};

	return {
		get fontFaceCss() {
			return embeddedCss;
		},
		get fontFamilies() {
			return families;
		},
		setExtraCss(css) {
			extraCss = css;
			paint();
		},
		set(fonts) {
			// The same list again changes nothing (the store re-emits on other fields).
			if (fonts === lastFonts) {
				return;
			}
			lastFonts = fonts;
			const previous = liveUrls;
			const { fontFaceCss, fontFamilies, objectUrls } = buildEmbeddedFontStyles(
				fonts,
				mintObjectUrl,
			);
			liveUrls = objectUrls;
			embeddedCss = fontFaceCss;
			families = fontFamilies;
			paint();
			revoke(previous);
		},
		dispose() {
			remove();
			revoke(liveUrls);
			liveUrls = [];
			lastFonts = undefined;
			embeddedCss = '';
			extraCss = '';
			families = [];
		},
	};
}
