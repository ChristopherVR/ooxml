/**
 * viewer-deck-style.service.ts: the Design galleries' hover preview and their effect and
 * background commits over the Angular editor state.
 *
 * Shared (`createDeckStyleHost`) decides what a pick does to the slides, the template elements and
 * the theme; this reads and writes the editor and loader signals, raises
 * {@link EditorStateService.deckStyleHeld} while a preview shows (collaboration then keeps to the
 * real slides) and records one undo step when a pick commits.
 */
import { inject, Injectable, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { createDeckStyleHost } from 'ooxml-ui/pptx';
import type { DeckStyleHandler, DeckStyleHost, DeckStyleState } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';

import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';

@Injectable()
export class ViewerDeckStyleService {
	private readonly editor = inject(EditorStateService);
	private readonly loader = inject(LoadContentService, { optional: true });
	private readonly translate = inject(TranslateService, { optional: true });

	/** The slide a "selected slides" pick lands on; the gallery sets it from its input. */
	readonly activeSlideIndex = signal(0);

	/** The deck a save or recovery snapshot writes: never a Design gallery hover preview. */
	savedDeck(): Pick<DeckStyleState, 'slides' | 'templateElementsBySlideId'> {
		const held = this.host.held();
		return {
			slides: held?.slides ?? (this.editor.slides() as PptxSlide[]),
			templateElementsBySlideId:
				held?.templateElementsBySlideId ?? this.editor.templateElementsBySlideId(),
		};
	}

	private write(state: DeckStyleState): void {
		this.editor.slides.set(state.slides);
		this.editor.templateElementsBySlideId.set(state.templateElementsBySlideId);
		this.loader?.theme.set(state.theme);
		this.loader?.themeColorMap.set(state.themeColorMap);
	}

	readonly host: DeckStyleHost = createDeckStyleHost(() => ({
		read: () => ({
			slides: this.editor.slides() as PptxSlide[],
			templateElementsBySlideId: this.editor.templateElementsBySlideId(),
			theme: this.loader?.theme(),
			themeColorMap: this.loader?.themeColorMap(),
		}),
		show: (state) => this.write(state),
		hold: (on) => this.editor.deckStyleHeld.set(on),
		handler: () => (this.loader?.getHandler() ?? null) as unknown as DeckStyleHandler | null,
		selectedSlideIds: () => {
			const slide = this.editor.slides()[this.activeSlideIndex()];
			return slide ? [slide.id] : [];
		},
		apply: (state) => {
			// Records the undo step from the restored slides, then takes the new ones.
			this.editor.applyReplacement(
				state.slides,
				this.translate?.instant('pptx.undoAction.applyTheme', { name: state.theme?.name ?? '' }) ??
					'Apply theme',
			);
			this.editor.templateElementsBySlideId.set(state.templateElementsBySlideId);
			this.loader?.theme.set(state.theme);
			this.loader?.themeColorMap.set(state.themeColorMap);
		},
	}));
}
