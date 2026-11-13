import { Directive, inject, Injectable, input, TemplateRef, ViewContainerRef } from '@angular/core';
import type { DoCheck, EmbeddedViewRef } from '@angular/core';
import {
	elementRenderErrorContext,
	elementRenderRetryKey,
	mergeElementRenderErrorScope,
	reportElementRenderError,
	shouldRetryElementRender,
} from 'ooxml-ui/pptx';
import type {
	ElementRenderErrorHandler,
	ElementRenderErrorScope,
	ElementRenderRetryKey,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

/**
 * The element render error scope of a subtree: the host's `onElementRenderError` (provided by
 * `pptx-viewer` and the building-blocks canvas and slides pane) and the painted slide's id
 * (provided by every `pptx-slide-canvas`, so thumbnails and the show report their own slide).
 * Provide it in a component's `providers` and set the getters in its constructor; an unset
 * getter inherits the enclosing scope.
 */
@Injectable()
export class ElementRenderErrors {
	private readonly parent = inject(ElementRenderErrors, { optional: true, skipSelf: true });
	reporter: () => ElementRenderErrorHandler | undefined = () => undefined;
	slideId: () => string | undefined = () => undefined;

	current(): ElementRenderErrorScope {
		return mergeElementRenderErrorScope(this.parent?.current(), {
			slideId: this.slideId(),
			onElementRenderError: this.reporter(),
		});
	}
}

/**
 * `*pptxElementBoundary="element"`: omits one slide element whose creation, template or
 * lifecycle hook throws, so its neighbours, the slide and the viewer stay. Angular has no error
 * boundaries, so this is the element host's try/catch: the directive creates the element's view
 * and checks it inside a try on every check of the enclosing view, before that view checks it
 * again (a cheap no-op then). On a failure the view is destroyed and the host is told through the
 * scope's `onElementRenderError` (default `console.error`). A replacement element object or a
 * changed id renders it again; checking the same failed object does not.
 *
 * Not caught: event handler errors (as in React) and an error thrown when only a signal inside
 * the element changed and Angular refreshes that view alone; those reach the `ErrorHandler`.
 */
@Directive({ selector: '[pptxElementBoundary]', standalone: true })
export class ElementBoundaryDirective implements DoCheck {
	readonly pptxElementBoundary = input.required<PptxElement>();
	private readonly template = inject<TemplateRef<unknown>>(TemplateRef);
	private readonly container = inject(ViewContainerRef);
	private readonly scope = inject(ElementRenderErrors, { optional: true });
	private view: EmbeddedViewRef<unknown> | null = null;
	private failed: ElementRenderRetryKey | null = null;

	ngDoCheck(): void {
		const element = this.pptxElementBoundary();
		if (!shouldRetryElementRender(this.failed, element)) {
			return;
		}
		this.failed = null;
		try {
			this.view ??= this.container.createEmbeddedView(this.template);
			this.view.detectChanges();
		} catch (error) {
			this.failed = elementRenderRetryKey(element);
			this.discard();
			const { slideId, onElementRenderError } = this.scope?.current() ?? {};
			reportElementRenderError(
				onElementRenderError,
				error,
				elementRenderErrorContext(element, slideId),
			);
		}
	}

	private discard(): void {
		this.view = null;
		try {
			this.container.clear();
		} catch {
			// A view that failed half way may not tear down cleanly; it is gone either way.
		}
	}
}
