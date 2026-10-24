/**
 * viewer-theme-gallery.service.ts: Viewer-scoped state + logic for the theme
 * gallery overlay (Design ▸ Browse Themes): its open/closed signal, the loaded
 * deck's active theme name (for the gallery's check-mark), and applying a
 * built-in preset to the whole deck.
 *
 * Extracted from {@link PowerPointViewerComponent}. Unlike most of the other
 * extracted viewer-*.service.ts controllers this one needs no `bind()`: it
 * only depends on other already-provided services (`EditorStateService`,
 * `LoadContentService`), so it injects them directly.
 *
 * Provide it once on the viewer component (`providers: [ViewerThemeGalleryService]`).
 */

import { computed, inject, Injectable, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { historyTheme, loadedThemeName, queueThemeWrite, writeThemeParts } from 'ooxml-ui/pptx';
import type { ThemeEditorEdit } from 'ooxml-ui/pptx';
import { applyThemeToData, reResolveElementColors, reResolveElementFonts } from 'pptx-viewer-core';
import type {
	PptxData,
	PptxThemeColorScheme,
	PptxThemeFontScheme,
	PptxThemePreset,
} from 'pptx-viewer-core';

import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import type { TemplateElementsBySlideId } from './template-mode';

@Injectable()
export class ViewerThemeGalleryService {
	private readonly editor = inject(EditorStateService);
	private readonly loader = inject(LoadContentService);
	private readonly translate = inject(TranslateService);

	/** Whether the theme-gallery overlay is visible (Design → Browse Themes). */
	readonly showThemeGallery = signal(false);
	readonly themeEditorApplying = signal(false);
	async applyThemeEditor(edit: ThemeEditorEdit): Promise<void> {
		if (this.themeEditorApplying()) {
			return;
		}
		this.themeEditorApplying.set(true);
		try {
			// On the handler's theme write queue: an undo still writing a theme part, or a save, waits
			// for it rather than interleaving with it.
			const handler = this.loader.getHandler();
			if (handler) {
				await queueThemeWrite(handler, () =>
					handler.applyTheme(edit.colorScheme, edit.fontScheme, edit.name),
				);
			}
			this.reTheme(edit.colorScheme, edit.fontScheme, edit.name, edit.name);
			this.showThemeGallery.set(false);
		} finally {
			this.themeEditorApplying.set(false);
		}
	}
	/** The `name` property of the loaded deck's theme (for check-mark in gallery). */
	readonly activeThemeName = computed<string | undefined>(() => this.loader.theme()?.name);

	/**
	 * Apply a built-in theme preset to the whole deck.
	 *
	 * Mirrors Vue's `applyThemePreset()`: re-resolves slide colours via core's
	 * pure `applyThemeToData`, then writes the updated slides + theme metadata
	 * into `EditorStateService` as a single undoable entry and the theme into the
	 * handler's theme parts. Also refreshes the `loader.themeColorMap` so
	 * subsequent theme switches start from the correct baseline.
	 */
	applyThemePreset(preset: PptxThemePreset): void {
		this.applyCustomTheme(preset.colorScheme, preset.fontScheme, preset.name);
	}

	applyCustomTheme(
		colorScheme: PptxThemeColorScheme,
		fontScheme: PptxThemeFontScheme,
		name: string,
	): void {
		void this.applySchemes(colorScheme, fontScheme, name, name);
		this.showThemeGallery.set(false);
	}

	/**
	 * Design > Variants (Colors / Fonts galleries): swap ONE scheme of the
	 * current theme and keep the other, without renaming the theme. Same
	 * undoable path as {@link applyCustomTheme}; `name` only labels the undo
	 * entry. A colour-less theme with a fonts pick has nothing to re-resolve
	 * against, so it is a no-op.
	 */
	applyThemeVariant(
		variant: { colorScheme?: PptxThemeColorScheme; fontScheme?: PptxThemeFontScheme },
		name: string,
	): void {
		const colorScheme = variant.colorScheme ?? this.loader.theme()?.colorScheme;
		if (!colorScheme) {
			return;
		}
		void this.applySchemes(colorScheme, variant.fontScheme, undefined, name);
	}

	/**
	 * {@link reTheme}, then write the parts it changed into the handler's theme parts (only those:
	 * a fonts-only change leaves the colour scheme part as authored). Resolves once written.
	 */
	private applySchemes(
		colorScheme: PptxThemeColorScheme,
		fontScheme: PptxThemeFontScheme | undefined,
		themeName: string | undefined,
		label: string,
	): Promise<void> {
		const from = historyTheme(this.loader.theme());
		this.reTheme(colorScheme, fontScheme, themeName, label);
		return writeThemeParts(
			this.loader.getHandler(),
			from,
			historyTheme(this.loader.theme()) ?? {},
			loadedThemeName(this.loader.themeOptions(), this.loader.slideMasters()),
		);
	}

	/** The slides, the template layer, the theme and its colour map, as one undo step. */
	private reTheme(
		colorScheme: PptxThemeColorScheme,
		fontScheme: PptxThemeFontScheme | undefined,
		themeName: string | undefined,
		label: string,
	): void {
		const currentSlides = this.editor.slides();
		const previousColorMap = this.loader.themeColorMap() ?? {};
		const previousFonts = this.loader.theme()?.fontScheme;
		const result = applyThemeToData(
			{
				slides: [...currentSlides],
				theme: this.loader.theme() ?? {},
				themeColorMap: previousColorMap,
			} as unknown as PptxData,
			colorScheme,
			fontScheme,
			themeName,
		);
		// Write slides back through the editor (records undo history).
		this.editor.applyReplacement(
			result.slides,
			this.translate.instant('pptx.undoAction.applyTheme', { name: label }),
		);
		// Master/layout elements render as a separate per-slide layer (not part
		// of `slide.elements`), so `applyReplacement` above never touches them;
		// left alone they'd keep painting the old scheme's colours.
		const currentTemplateElements = this.editor.templateElementsBySlideId();
		if (Object.keys(currentTemplateElements).length > 0) {
			const recoloured: TemplateElementsBySlideId = {};
			for (const [slideId, elements] of Object.entries(currentTemplateElements)) {
				recoloured[slideId] = reResolveElementFonts(
					reResolveElementColors(elements, previousColorMap, colorScheme),
					previousFonts,
					fontScheme ?? previousFonts ?? {},
				);
			}
			this.editor.templateElementsBySlideId.set(recoloured);
		}
		// Update the loader's theme signals so the check-mark and future switches are correct.
		this.loader.theme.set(result.theme);
		this.loader.themeColorMap.set(result.themeColorMap);
	}
}
