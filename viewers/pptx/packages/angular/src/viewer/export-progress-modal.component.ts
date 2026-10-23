/**
 * export-progress-modal.component.ts: progress overlay shown while a multi-slide
 * export (PDF / GIF / WebM) runs.
 *
 * Selector: `pptx-export-progress-modal`
 *
 * A thin adapter around the shared `pptx-ui-export-progress` element, which owns
 * the dialog and progress bar semantics, focus and the look (from the viewer's
 * theme). It is deliberately NOT dismissable by backdrop click or Escape - an
 * export in flight should only end by completing, erroring, or the user pressing
 * Cancel (which emits `cancel`; the host aborts the `AbortController` the export
 * loop checks between slides and clears `open`).
 *
 * Usage:
 * ```html
 * <pptx-export-progress-modal
 *   [open]="exportModalOpen()"
 *   [title]="exportModalTitle()"
 *   [progress]="exportProgress()"
 *   [statusMessage]="exportStatusMessage()"
 *   (cancel)="onCancelExport()"
 * />
 * ```
 */

import {
	ChangeDetectionStrategy,
	Component,
	computed,
	CUSTOM_ELEMENTS_SCHEMA,
	inject,
	input,
	output,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import type { ExportProgressViewState } from 'ooxml-ui/pptx';

import { translationsSignal } from './translations-signal';

@Component({
	selector: 'pptx-export-progress-modal',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	host: { class: 'contents' },
	template: `<pptx-ui-export-progress [state]="view()" (export-cancel)="cancel.emit()" />`,
})
export class ExportProgressModalComponent {
	/** Whether the overlay is visible. */
	readonly open = input<boolean>(false);
	/** Heading shown at the top (e.g. "Export as PDF"). */
	readonly title = input<string>('');
	/** Current progress, 0-100. */
	readonly progress = input<number>(0);
	/** Optional status line (e.g. "Rendering slide 3 of 10..."). */
	readonly statusMessage = input<string>('');

	/** Emitted when the user presses Cancel. */
	readonly cancel = output<void>();

	private readonly translate = inject(TranslateService);
	private readonly translations = translationsSignal(this.translate);

	protected readonly view = computed<ExportProgressViewState>(() => {
		this.translations();
		return {
			open: this.open(),
			title: this.title(),
			progress: this.progress(),
			status: this.statusMessage(),
			translate: (key, params) => this.translate.instant(key, params),
		};
	});
}
