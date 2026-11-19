/**
 * external-content-bar.component.ts: PowerPoint's security bar for linked
 * pictures and media.
 *
 * Selector: `pptx-external-content-bar`
 *
 * A thin adapter around the shared `pptx-ui-external-content-bar` element,
 * which owns the markup and the strings ("SECURITY WARNING ... Enable
 * Content"). The viewer decides when it shows (`ExternalContentService`)
 * and what Enable Content does; this component only translates and routes
 * the element's `external-content-request` through the shared
 * `dispatchExternalContentBarIntent`.
 *
 * @module viewer/external-content-bar
 */
import {
	ChangeDetectionStrategy,
	Component,
	computed,
	CUSTOM_ELEMENTS_SCHEMA,
	inject,
	output,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { dispatchExternalContentBarIntent } from 'ooxml-ui/pptx';
import type { ExternalContentBarViewState } from 'ooxml-ui/pptx';

import { translationsSignal } from './translations-signal';

@Component({
	selector: 'pptx-external-content-bar',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	host: { class: 'contents' },
	template: `
		<pptx-ui-external-content-bar [state]="view()" (external-content-request)="request($event)" />
	`,
})
export class ExternalContentBarComponent {
	/** Enable Content: draw the presentation's linked content. */
	readonly enable = output<void>();
	/** The close button: hide the bar, keep the content blocked. */
	readonly closeBar = output<void>();

	private readonly translate = inject(TranslateService);
	private readonly translations = translationsSignal(this.translate);

	protected readonly view = computed<ExternalContentBarViewState>(() => {
		this.translations();
		return { translate: (key, params) => this.translate.instant(key, params) };
	});

	protected request(event: Event): void {
		dispatchExternalContentBarIntent(event, {
			enable: () => this.enable.emit(),
			close: () => this.closeBar.emit(),
		});
	}
}
