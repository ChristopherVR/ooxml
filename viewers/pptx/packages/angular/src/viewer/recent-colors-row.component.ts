/**
 * recent-colors-row.component.ts: the reusable "Recent colours" swatch row
 * shown under a colour picker once at least one colour has been used
 * (`p:clrMru`, backed by {@link RecentColorsService}), drawn by the shared
 * `pptx-ui-color-grid` with only its Recent Colors section
 * (`recentColorsGridView`).
 *
 * Presentational only: clicking a swatch just emits `pick`. The caller both
 * applies the colour through its own commit path AND pushes it back into
 * `RecentColorsService` (which moves it to the front of the list), exactly
 * like React's `ColorPickerRow` and Vue's `RecentColorsRow.vue`.
 *
 * The inspector's fill, stroke and text colour pickers all mount this one row
 * instead of re-implementing it.
 *
 * @module viewer/recent-colors-row
 */
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
import type { OfficeColorPickEvent } from 'ooxml-ui/pptx';
import { recentColorsGridView } from 'ooxml-ui/pptx';

import { homeLanguage, ngxTranslate } from './ribbon-home-lang';

@Component({
	selector: 'pptx-recent-colors-row',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `
		@if (colors().length > 0) {
			<pptx-ui-color-grid
				class="pptx-recent-colors"
				data-testid="pptx-color-recent"
				[palette]="view().palette"
				[standardColors]="view().standardColors"
				[recentColors]="view().recentColors"
				[recentHeading]="view().recentHeading"
				[recentLabel]="view().recentLabel"
				[label]="view().label"
				[disabled]="disabled()"
				(mousedown)="$event.preventDefault()"
				(office-color-pick)="onPick($event)"
			></pptx-ui-color-grid>
		}
	`,
	styles: `
		.pptx-recent-colors {
			margin-top: 0.25rem;
		}
	`,
})
export class RecentColorsRowComponent {
	private readonly translation = inject(TranslateService, { optional: true });
	/** Read so a language switch re-labels the row. */
	private readonly language = homeLanguage(this.translation);

	/** Most-recently-used colours, most-recent-first, `#RRGGBB` uppercase. */
	readonly colors = input<readonly string[]>([]);
	/** Disable every swatch (mirrors the host picker's own disabled state). */
	readonly disabled = input<boolean>(false);
	/** A swatch was clicked; the caller both applies AND pushes the colour. */
	readonly pick = output<string>();

	protected readonly view = computed(() => {
		this.language();
		return recentColorsGridView(this.colors(), ngxTranslate(this.translation));
	});

	protected onPick(event: Event): void {
		this.pick.emit((event as OfficeColorPickEvent).detail.color);
	}
}
