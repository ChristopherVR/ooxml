/**
 * `EmbeddedFontsService`: Angular port of the Vue `useEmbeddedFonts`
 * composable (the `@font-face` half of the React `useFontInjection` hook).
 *
 * Given the embedded fonts parsed out of a `.pptx` by `pptx-viewer-core`, this
 * service builds `@font-face` CSS rules and injects them into a single managed
 * `<style>` element appended to `<head>`, so slides render with the typefaces
 * the author embedded. The pure string-building / format-mapping logic lives in
 * `./embedded-fonts-helpers`; this service owns only the DOM side effects and
 * the signal surface.
 *
 * Lifecycle mirrors `LoadContentService`: any object URLs minted from
 * de-obfuscated font bytes are tracked and revoked when superseded, and the
 * injected `<style>` plus all live object URLs are torn down on destroy via
 * `DestroyRef.onDestroy` (or an explicit {@link dispose}).
 *
 * Provide at the component level so its lifetime tracks the host viewer:
 * `@Component({ providers: [EmbeddedFontsService] })`.
 */

import { DestroyRef, Injectable, inject, signal } from '@angular/core';
import { createEmbeddedFontsSheet, createEmbeddedOutlineFontsSlot } from 'ooxml-ui/pptx';
import type { EmbeddedFontsSheet } from 'ooxml-ui/pptx';
import type { PptxEmbeddedFont } from 'pptx-viewer-core';

import { EMBEDDED_FONTS_STYLE_ID, buildUserFontFaceStyles } from './embedded-fonts-helpers';
import type { ViewerFontSource } from './embedded-fonts-helpers';
import { glyphOutlineFontCache } from './glyph-outline-cache';

/** True when the runtime exposes the DOM APIs we need to inject a `<style>`. */
function hasDomSupport(): boolean {
	return (
		typeof document !== 'undefined' &&
		typeof document.createElement === 'function' &&
		typeof document.head !== 'undefined'
	);
}

@Injectable()
export class EmbeddedFontsService {
	/** The generated `@font-face` stylesheet text (also injected into `<head>`). */
	readonly fontFaceCss = signal('');
	/**
	 * Distinct CSS `font-family` strings for the embedded families, each with
	 * substitution fallbacks resolved (e.g. `'"Calibri", "Carlito", …'`).
	 */
	readonly fontFamilies = signal<string[]>([]);

	/** The shared managed `<style>` element; owns the rule text and object URLs. */
	private readonly sheet: EmbeddedFontsSheet | null = hasDomSupport()
		? createEmbeddedFontsSheet(document, EMBEDDED_FONTS_STYLE_ID)
		: null;
	/** The current deck's fonts in the module-wide glyph-outline cache. */
	private readonly outlineFonts = createEmbeddedOutlineFontsSlot(glyphOutlineFontCache);

	constructor() {
		inject(DestroyRef).onDestroy(() => {
			this.dispose();
		});
	}

	/**
	 * Resolve the supplied embedded fonts into `@font-face` rules, inject them
	 * into the managed `<style>` element, and update the exposed signals.
	 * Pass an empty list (or `null`) to clear everything.
	 */
	setFonts(fonts: readonly PptxEmbeddedFont[] | null | undefined): void {
		// Register this deck's embedded fonts' real outlines (see
		// glyph-outline-cache.ts) for WordArt envelope glyph-outline warping.
		this.outlineFonts.set(fonts);
		if (!this.sheet) {
			return;
		}
		this.sheet.set(fonts);
		this.fontFaceCss.set(this.sheet.fontFaceCss);
		this.fontFamilies.set([...this.sheet.fontFamilies]);
	}

	/** Replace the licensed font sources supplied by the host application. */
	setHostFonts(fonts: readonly ViewerFontSource[] | null | undefined): void {
		this.sheet?.setExtraCss(buildUserFontFaceStyles(fonts ?? []));
	}

	/**
	 * Remove the injected `<style>` element, revoke all live object URLs, and
	 * reset the signals. Called automatically on destroy; safe to call manually.
	 */
	dispose(): void {
		this.outlineFonts.release();
		this.sheet?.dispose();
		this.fontFaceCss.set('');
		this.fontFamilies.set([]);
	}
}
