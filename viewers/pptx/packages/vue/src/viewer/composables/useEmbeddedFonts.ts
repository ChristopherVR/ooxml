import { createEmbeddedFontsSheet, createEmbeddedOutlineFontsSlot } from 'ooxml-ui/pptx';
import type { PptxEmbeddedFont } from 'pptx-viewer-core';
import { computed, onScopeDispose, shallowRef, toValue, watch } from 'vue';
import type { ComputedRef, MaybeRefOrGetter } from 'vue';

import { glyphOutlineFontCache } from '../utils/glyph-outline-cache';

/**
 * `useEmbeddedFonts`: Vue port of the embedded-font injection logic in the
 * React `useFontInjection` hook (the `@font-face` half).
 *
 * Given the embedded fonts parsed out of a `.pptx` by `pptx-viewer-core`, this
 * composable builds `@font-face` CSS rules and injects them into the document
 * so that slides render with the typefaces the author embedded. The injected
 * `<style>` element is appended to `<head>` and removed on scope dispose.
 *
 * The core load pipeline (`PptxHandlerRuntimeEmbeddedFonts`) already
 * de-obfuscates the OOXML XOR layer and produces a ready-to-use `dataUrl` for
 * each style variant. One {@link PptxEmbeddedFont} therefore corresponds to a
 * single style variant (regular / bold / italic / boldItalic), distinguished by
 * its `bold` and `italic` flags: a font with regular + bold data yields two
 * entries and two `@font-face` rules.
 *
 * The rule text, the de-obfuscation fallback and the object-URL lifecycle are
 * the shared `createEmbeddedFontsSheet`; this composable only owns the watch.
 *
 * @param fonts - Reactive source of the parsed embedded fonts (the host feeds
 *   these from `PptxData.embeddedFonts`).
 * @returns Reactive `fontFaceCss` (the generated stylesheet text) and
 *   `fontFamilies` (the distinct embedded family names, with substitution
 *   fallbacks resolved for downstream `font-family` use).
 */
export interface UseEmbeddedFontsResult {
	/** The generated `@font-face` stylesheet text (also injected into `<head>`). */
	fontFaceCss: ComputedRef<string>;
	/**
	 * Distinct CSS `font-family` strings for the embedded families, each with
	 * substitution fallbacks resolved (e.g. `'"Calibri", "Carlito", …'`).
	 */
	fontFamilies: ComputedRef<string[]>;
}

const STYLE_ELEMENT_ID = 'pptx-vue-embedded-fonts';

export function useEmbeddedFonts(
	fonts: MaybeRefOrGetter<PptxEmbeddedFont[]>,
): UseEmbeddedFontsResult {
	// The shared sheet owns the <style> element, the rule text and the object-URL
	// lifecycle (revoked when a source is superseded and on dispose).
	const sheet =
		typeof document === 'undefined' || typeof document.head === 'undefined'
			? null
			: createEmbeddedFontsSheet(document, STYLE_ELEMENT_ID);
	const cssRef = shallowRef('');
	const familiesRef = shallowRef<string[]>([]);

	const outlineFonts = createEmbeddedOutlineFontsSlot(glyphOutlineFontCache);
	// `immediate: true` registers the outlines synchronously at composable setup,
	// before the first render that would use them (see glyph-outline-cache.ts).
	const stopWatch = watch(
		() => toValue(fonts) ?? [],
		(source) => {
			outlineFonts.set(source);
			sheet?.set(source);
			if (sheet) {
				cssRef.value = sheet.fontFaceCss;
				familiesRef.value = [...sheet.fontFamilies];
			}
		},
		{ immediate: true },
	);

	onScopeDispose(() => {
		stopWatch();
		outlineFonts.release();
		sheet?.dispose();
	});

	return {
		fontFaceCss: computed(() => cssRef.value),
		fontFamilies: computed(() => familiesRef.value),
	};
}
