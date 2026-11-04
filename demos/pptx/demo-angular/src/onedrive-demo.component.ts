import {
	ChangeDetectionStrategy,
	Component,
	DestroyRef,
	type ElementRef,
	afterNextRender,
	inject,
	signal,
	viewChild,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { PowerPointViewerComponent, translationsEn } from 'pptx-angular-viewer';

import { mountOneDriveDemo } from '../../shared/onedrive-demo';
import type { OneDriveDemoDeck, OneDriveDemoMount } from '../../shared/onedrive-demo';

/** `?onedrive=1`: open, edit and save a deck on your own OneDrive (docs/onedrive-quickstart.md). */
@Component({
	selector: 'app-root',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [PowerPointViewerComponent],
	template: `<div #bar></div>
		<main class="od-viewer">
			@if (deck(); as current) {
				<pptx-viewer
					[content]="current.bytes"
					[fileName]="current.fileName"
					[canEdit]="current.canEdit"
					[storage]="storage()"
				/>
			} @else {
				<p class="od-empty">No OneDrive file open yet.</p>
			}
		</main>`,
})
export class OneDriveDemoComponent {
	readonly deck = signal<OneDriveDemoDeck | null>(null);
	readonly storage = signal<OneDriveDemoMount['storage'] | undefined>(undefined);
	private readonly bar = viewChild.required<ElementRef<HTMLDivElement>>('bar');
	constructor() {
		const translate = inject(TranslateService);
		translate.setTranslation('en', translationsEn);
		translate.use('en');
		const destroyRef = inject(DestroyRef);
		afterNextRender(() => {
			const mounted = mountOneDriveDemo(this.bar().nativeElement, (deck) => this.deck.set(deck));
			this.storage.set(mounted.storage);
			destroyRef.onDestroy(() => mounted.dispose());
		});
	}
}
