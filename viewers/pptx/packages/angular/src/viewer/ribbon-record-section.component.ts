import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	computed,
	inject,
	input,
	output,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { buildRecordRibbon, runRecordRibbonAction } from 'ooxml-ui/pptx';
import type { RecordTarget, RibbonCommandRequestEvent } from 'ooxml-ui/pptx';

import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { translationsSignal } from './translations-signal';

@Component({
	selector: 'pptx-ribbon-record-section',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { class: 'contents' },
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `<pptx-ui-ribbon-section [groups]="groups()" (command-request)="request($event)" />`,
})
export class RibbonRecordSectionComponent {
	private readonly translate = inject(TranslateService);
	private readonly editor = inject(EditorStateService, { optional: true });
	private readonly loader = inject(LoadContentService, { optional: true });
	private readonly translations = translationsSignal(this.translate);
	readonly canEdit = input(false);
	readonly slideIndex = input(0);
	readonly recordFromBeginning = output<void>();
	readonly recordFromCurrent = output<void>();
	protected readonly groups = computed(() => {
		this.translations();
		return buildRecordRibbon(this.t, this.target());
	});
	private readonly t = (key: string): string => this.translate.instant(key) as string;
	private readonly target = (): RecordTarget | undefined => {
		const { editor, loader } = this;
		if (!this.canEdit() || !editor || !loader) return undefined;
		const slideIndex = this.slideIndex();
		return {
			slides: editor.slides(),
			slideIndex,
			canvasSize: loader.canvasSize(),
			insert: (element) => editor.addElement(slideIndex, element),
			commitSlides: (slides) => editor.applyReplacement(slides, this.t('pptx.ribbon.tab.record')),
		};
	};
	protected request(event: Event): void {
		const id = (event as RibbonCommandRequestEvent).detail.id;
		if (id === 'record.record.fromBeginning') {
			this.recordFromBeginning.emit();
		} else if (id === 'record.record.fromCurrent') {
			this.recordFromCurrent.emit();
		} else if (event.currentTarget instanceof HTMLElement)
			runRecordRibbonAction(event.currentTarget, id, this.target, this.t);
	}
}
