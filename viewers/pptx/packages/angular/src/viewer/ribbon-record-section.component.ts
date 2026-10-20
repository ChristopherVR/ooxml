import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	computed,
	inject,
	output,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { buildRecordRibbon } from 'ooxml-ui/pptx';
import type { RibbonCommandRequestEvent } from 'ooxml-ui/pptx';

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
	private readonly translations = translationsSignal(this.translate);
	readonly recordFromBeginning = output<void>();
	readonly recordFromCurrent = output<void>();
	protected readonly groups = computed(() => {
		this.translations();
		return buildRecordRibbon((key) => this.translate.instant(key) as string);
	});
	protected request(event: Event): void {
		const id = (event as RibbonCommandRequestEvent).detail.id;
		if (id === 'record.record.fromBeginning') {
			this.recordFromBeginning.emit();
		} else if (id === 'record.record.fromCurrent') {
			this.recordFromCurrent.emit();
		}
	}
}
