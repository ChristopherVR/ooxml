import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	computed,
	inject,
	output,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { buildHelpRibbon } from 'ooxml-ui/pptx';
import type { RibbonCommandRequestEvent } from 'ooxml-ui/pptx';

import { translationsSignal } from './translations-signal';
import { ViewerCustomizationService } from './viewer-customization.service';

@Component({
	selector: 'pptx-ribbon-help-section',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { class: 'contents' },
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `<pptx-ui-ribbon-section [groups]="groups()" (command-request)="request($event)" />`,
})
export class RibbonHelpSectionComponent {
	private readonly translate = inject(TranslateService);
	private readonly translations = translationsSignal(this.translate);
	private readonly customization = inject(ViewerCustomizationService, { optional: true });
	readonly openSettings = output<void>();
	readonly openShortcuts = output<void>();
	readonly a11y = output<void>();
	protected readonly groups = computed(() => {
		this.translations();
		return buildHelpRibbon((key) => this.translate.instant(key) as string, {
			options: this.customization?.dialogAvailable('options') !== false,
		});
	});
	protected request(event: Event): void {
		switch ((event as RibbonCommandRequestEvent).detail.id) {
			case 'help.help.options':
				this.openSettings.emit();
				break;
			case 'help.help.keyboardShortcuts':
				this.openShortcuts.emit();
				break;
			case 'help.help.accessibility':
				this.a11y.emit();
				break;
		}
	}
}
