import { ChangeDetectionStrategy, Component, DestroyRef, inject, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { PowerPointViewerComponent, translationsEn } from 'pptx-angular-viewer';
import {
	translationsDe,
	translationsEs,
	translationsFr,
	translationsZhCN,
} from 'pptx-viewer-locales';

import { createSameBrowserDemo } from '../../shared/same-browser-collaboration';
import type { SameBrowserDemo } from '../../shared/same-browser-collaboration';

@Component({
	selector: 'app-root',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [PowerPointViewerComponent],
	template: `<main style="position: fixed; inset: 28px 0 0">
		@if (error()) {
			<p role="alert">{{ error() }}</p>
		}
		@if (demo(); as session) {
			<pptx-viewer
				[content]="session.source"
				[fileName]="session.fileName"
				[collaboration]="session.collaboration"
				[canEdit]="true"
				(collaborationChange)="session.report($event)"
			/>
		}
	</main>`,
})
export class SameBrowserDemoComponent {
	readonly demo = signal<SameBrowserDemo | null>(null);
	readonly error = signal('');
	constructor() {
		const translate = inject(TranslateService);
		for (const [language, messages] of Object.entries({
			en: translationsEn,
			de: translationsDe,
			es: translationsEs,
			fr: translationsFr,
			'zh-CN': translationsZhCN,
		})) {
			translate.setTranslation(language, messages);
		}
		translate.use('en');
		let disposed = false;
		inject(DestroyRef).onDestroy(() => {
			disposed = true;
			this.demo()?.dispose();
		});
		void createSameBrowserDemo()
			.then((value) => {
				if (disposed) {
					value.dispose();
					return;
				}
				this.demo.set(value);
				return undefined;
			})
			.catch((reason: unknown) => this.error.set(String(reason)));
	}
}
