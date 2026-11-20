/**
 * external-content.service.ts: PowerPoint's security bar for linked web content, per viewer.
 *
 * The deck model always keeps its linked web addresses (the loader parses with
 * `allowExternalImages: true`), so they are saved, collaborated and undone like any other
 * content. What the browser may fetch is this viewer's choice: the shared
 * `resolveExternalContentView` derives it from the host `allowExternalContent` input, the Trust
 * Center option and the consent the loader keeps for the presentation now open. Every surface
 * that draws the model reads it through {@link LINKED_CONTENT_ALLOWED} and draws
 * `linkedContentView(value, allowed)`, so Enable Content only redraws: nothing reloads, and the
 * deck, its undo history, autosave and a collaboration session carry on.
 *
 * @module viewer/external-content
 */
import { computed, inject, Injectable, InjectionToken, signal } from '@angular/core';
import type { Signal } from '@angular/core';
import { resolveExternalContentView } from 'ooxml-ui/pptx';
import type { ExternalContentView } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { findOwningSlideIndex } from './smart-art-inline-edit';
import { ViewerOptionsService } from './viewer-options.service';

const NO_TRUST_OPTIONS = { trust: { allowExternalContent: false } };

@Injectable()
export class ExternalContentService {
	private readonly loader = inject(LoadContentService);
	private readonly editor = inject(EditorStateService);
	private readonly options = inject(ViewerOptionsService, { optional: true });

	/**
	 * What this viewer shows of the deck's linked web content. Everything it draws counts: a deck
	 * adopted from a collaboration room shows the bar as much as a loaded one.
	 */
	readonly view: Signal<ExternalContentView> = computed(() =>
		resolveExternalContentView({
			host: this.loader.hostAllowExternalContent(),
			options: this.options?.options() ?? NO_TRUST_OPTIONS,
			...this.loader.externalContentConsent(),
			content: [
				this.editor.slides(),
				this.editor.templateElementsBySlideId(),
				this.loader.tableStyleMap(),
			],
		}),
	);
	/** Surfaces draw the linked content itself. */
	readonly allowed = computed(() => this.view().allowed);
	/** The security bar shows. */
	readonly barVisible = computed(() => this.view().barVisible);

	/** Enable Content: draw this presentation's linked content, in this viewer only. */
	enable(): void {
		this.loader.externalContentConsent.update((consent) => ({ ...consent, enabled: true }));
	}

	/** The bar's close button: hide it, keep the content withheld. */
	close(): void {
		this.loader.externalContentConsent.update((consent) => ({ ...consent, closed: true }));
	}
}

/**
 * Whether the surfaces under this injector draw the deck's linked web content. The viewer
 * provides {@link ExternalContentService.allowed}; a `SlideCanvasComponent` with a
 * `linkedContent` input provides that value to what it draws.
 */
export const LINKED_CONTENT_ALLOWED = new InjectionToken<Signal<boolean>>(
	'pptx.linkedContentAllowed',
);

/** The provider every viewer and building-blocks host gets from `POWER_POINT_VIEWER_PROVIDERS`. */
export const LINKED_CONTENT_ALLOWED_PROVIDER = {
	provide: LINKED_CONTENT_ALLOWED,
	useFactory: (): Signal<boolean> => inject(ExternalContentService).allowed,
};

const WITHHELD = signal(false).asReadonly();

/** {@link LINKED_CONTENT_ALLOWED} here, or withheld outside any viewer (fail closed). */
export function injectLinkedContentAllowed(): Signal<boolean> {
	return inject(LINKED_CONTENT_ALLOWED, { optional: true }) ?? WITHHELD;
}

/**
 * The model element `id` names and its slide: what an on-canvas edit starts from, never the
 * drawn copy (which may withhold linked content). `templateSlideId` is the slide the hosting
 * canvas shows (`SLIDE_CONTEXT`), for template elements.
 */
export function modelElementById(
	editor: EditorStateService | null,
	id: string,
	templateSlideId: string | null,
): { slideIndex: number; element: PptxElement } | undefined {
	if (!editor) {
		return undefined;
	}
	const slideIndex = findOwningSlideIndex(editor.slides(), id, templateSlideId);
	const element = slideIndex < 0 ? undefined : editor.elementById(slideIndex, id);
	return element ? { slideIndex, element } : undefined;
}
