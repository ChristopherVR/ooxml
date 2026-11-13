import { Injectable, signal } from '@angular/core';
import type { Signal } from '@angular/core';
import type { ElementFindHighlights, SlideFindHighlights } from 'ooxml-ui/pptx';

/**
 * The Find bar's highlights on a slide canvas, by element id (shared
 * `findTextHighlightsByElement`), provided by each `SlideCanvasComponent` from
 * its `findHighlights` input, so the text renderers under it (and only those)
 * mark the matches. Mirrors Vue's `FindHighlightsKey` and Svelte's context.
 */
@Injectable()
export class CanvasFindHighlights {
	private readonly source = signal<Signal<SlideFindHighlights | undefined> | undefined>(undefined);

	/** Read the canvas's `findHighlights` input (called once by the canvas). */
	bind(source: Signal<SlideFindHighlights | undefined>): void {
		this.source.set(source);
	}

	/** The highlights of one element, read reactively. */
	forElement(elementId: string): ElementFindHighlights | undefined {
		return this.source()?.()?.get(elementId);
	}
}
