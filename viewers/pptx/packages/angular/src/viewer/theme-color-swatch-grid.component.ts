/**
 * theme-color-swatch-grid.component.ts: PowerPoint's "Theme Colors" grid (ten
 * columns - Background 1, Text 1, Background 2, Text 2, Accent 1..6 - each
 * with a base swatch and five luminance variants), built from the loaded
 * deck's real theme colours (`LoadContentService.themeColorMap`) rather than
 * a hard-coded Office palette, and drawn by the shared `pptx-ui-color-grid`.
 *
 * Mirrors React's `ThemeColorSwatchGrid.tsx` and Vue's
 * `ThemeColorSwatchGrid.vue`: a presentational grid that emits `pick` with
 * BOTH the resolved hex and the `PptxThemeColorRef` a caller should store so
 * the colour keeps following the theme after a later theme change.
 *
 * Renders nothing when no deck theme is loaded yet, so callers can render
 * this unconditionally alongside their existing hex/recent-colour controls.
 *
 * @module viewer/theme-color-swatch-grid
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
import type { OfficeColorPickEvent, ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import { buildThemeColorSwatchGrid, themeColorGridCommit, themeColorGridView } from 'ooxml-ui/pptx';
import type { PptxThemeColorRef } from 'pptx-viewer-core';

import { LoadContentService } from './load-content.service';
import { homeLanguage, ngxTranslate } from './ribbon-home-lang';

@Component({
	selector: 'pptx-theme-color-swatch-grid',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `
		@if (columns().length > 0) {
			<pptx-ui-color-grid
				class="pptx-theme-swatch-grid"
				[palette]="view().palette"
				[selectedTheme]="view().selectedTheme"
				[standardColors]="noStandardColors"
				[themeHeading]="heading()"
				[label]="heading()"
				[disabled]="disabled()"
				(mousedown)="$event.preventDefault()"
				(office-color-pick)="onPick($event)"
			></pptx-ui-color-grid>
		}
	`,
	styles: `
		.pptx-theme-swatch-grid {
			margin-top: 0.25rem;
		}
	`,
})
export class ThemeColorSwatchGridComponent {
	private readonly loader = inject(LoadContentService);
	private readonly translation = inject(TranslateService, { optional: true });
	/** Read so a language switch re-labels the swatches. */
	private readonly language = homeLanguage(this.translation);

	readonly disabled = input<boolean>(false);
	/** The element's current theme ref, if any (highlights the matching swatch). */
	readonly selectedRef = input<PptxThemeColorRef | undefined>(undefined);
	/** The element's current resolved hex, used to highlight a swatch when no ref is stored. */
	readonly selectedHex = input<string | undefined>(undefined);
	/** A swatch was clicked: both the resolved hex and the ref to store. */
	readonly pick = output<ThemeColorPickerCommit>();

	/** The inspector keeps its own Standard Colors and Recent Colors rows. */
	protected readonly noStandardColors = [];
	protected readonly columns = computed(() => {
		this.language();
		return buildThemeColorSwatchGrid(this.loader.themeColorMap(), ngxTranslate(this.translation));
	});
	protected readonly view = computed(() =>
		themeColorGridView(this.columns(), this.selectedRef(), this.selectedHex()),
	);
	protected readonly heading = computed(() => {
		this.language();
		return ngxTranslate(this.translation)('pptx.colorPicker.themeColors');
	});

	protected onPick(event: Event): void {
		const commit = themeColorGridCommit(this.columns(), (event as OfficeColorPickEvent).detail);
		if (commit) {
			this.pick.emit(commit);
		}
	}
}
