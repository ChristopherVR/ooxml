/**
 * list-conversion-scope.service.ts: the host's `listConversionScope` input, shared through DI
 * with every Bullets / Numbering call site (ribbon Paragraph controls, the inline editor's
 * format shortcuts, the master-view canvas) so the input is not threaded through each chrome
 * component.
 *
 * `PowerPointViewerComponent` binds its input signal once; call sites read {@link options}
 * at command time. Unbound (a custom host that never binds it) keeps the shared default,
 * `'paragraph'`.
 *
 * Provided per viewer (see `POWER_POINT_VIEWER_PROVIDERS`).
 */
import { Injectable, signal } from '@angular/core';
import type { Signal } from '@angular/core';
import type { ListConversionScope, ListStyleUpdateOptions } from 'ooxml-ui/pptx';

@Injectable()
export class ListConversionScopeService {
	private readonly source = signal<Signal<ListConversionScope | undefined> | null>(null);

	/** Follow the viewer's `listConversionScope` input. Returns the service, for a field initialiser. */
	bind(scope: Signal<ListConversionScope | undefined>): this {
		this.source.set(scope);
		return this;
	}

	/** The options to hand the shared list-style helpers for the current input value. */
	options(): ListStyleUpdateOptions {
		return { listConversionScope: this.source()?.() };
	}
}
