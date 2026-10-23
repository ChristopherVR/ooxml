/**
 * insert-dialogs-host.component.ts: hosts the shared PowerPoint Insert SmartArt
 * (`pptx-ui-insert-smartart-dialog`) and Insert Chart (`pptx-ui-insert-chart-dialog`)
 * dialogs from `ooxml-ui/pptx`. The elements own the view (pickers, previews,
 * focus, Escape); this component only maps their events onto the editor:
 * an insert adds the element through `EditorStateService.addElement` (one history
 * entry), and every close clears the host's open state.
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
import type { InsertChartDialogInsertEvent, InsertSmartArtDialogInsertEvent } from 'ooxml-ui/pptx';

import { newChartElement } from './editor-insert';
import { EditorStateService } from './editor-state.service';
import { homeLanguage, ngxTranslate } from './ribbon-home-lang';
import { buildSmartArtInsertElement } from './smart-art-insert-helpers';

@Component({
	selector: 'pptx-insert-dialogs-host',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { class: 'contents' },
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `
		<pptx-ui-insert-smartart-dialog
			[open]="smartArtOpen()"
			[translateLabel]="translate()"
			(insert-smartart)="onInsertSmartArt($event)"
			(dialog-close)="smartArtClosed.emit()"
		/>
		<pptx-ui-insert-chart-dialog
			[open]="chartOpen()"
			[translateLabel]="translate()"
			(insert-chart)="onInsertChart($event)"
			(dialog-close)="chartClosed.emit()"
		/>
	`,
})
export class InsertDialogsHostComponent {
	readonly smartArtOpen = input<boolean>(false);
	readonly chartOpen = input<boolean>(false);
	readonly slideIndex = input<number>(0);
	readonly smartArtClosed = output<void>();
	readonly chartClosed = output<void>();

	private readonly editor = inject(EditorStateService);
	private readonly translation = inject(TranslateService, { optional: true });
	private readonly language = homeLanguage(this.translation);
	protected readonly translate = computed(() => {
		this.language();
		return ngxTranslate(this.translation);
	});

	protected onInsertSmartArt(event: Event): void {
		const { layout, items } = (event as InsertSmartArtDialogInsertEvent).detail;
		this.editor.addElement(this.slideIndex(), buildSmartArtInsertElement(layout, items));
	}

	protected onInsertChart(event: Event): void {
		const { kind, title, grouping } = (event as InsertChartDialogInsertEvent).detail;
		this.editor.addElement(this.slideIndex(), newChartElement(kind, { title, grouping }));
	}
}
