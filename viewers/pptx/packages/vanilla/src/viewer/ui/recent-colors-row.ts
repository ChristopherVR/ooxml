import type { OfficeColorPickEvent, PptxUiColorGridElement } from 'ooxml-ui/pptx';
import { recentColorsGridView, registerPptxWebControls } from 'ooxml-ui/pptx';

import type { Translator } from '../i18n';

/**
 * B6: the reusable "Recent colours" (`p:clrMru`) swatch row, shown under a
 * colour input (fill, stroke, text) in the property inspector, drawn by the
 * shared `pptx-ui-color-grid` with only its Recent Colors section
 * (`recentColorsGridView`).
 *
 * Row contract (asserted by the cross-binding e2e spec): the grid carries
 * `data-testid="pptx-color-recent"` and an `aria-label` matching its visible
 * heading, one swatch per colour (most-recent first) named
 * "Recent <hex>", and the whole row hidden while the list is empty.
 */
export interface RecentColorsRowHandle {
	el: HTMLElement;
	/** Replace the shown colours (most-recent-first); hides the row when empty. */
	setColors(colors: readonly string[]): void;
	setDisabled(disabled: boolean): void;
}

/** Build a "Recent colours" row. `onSelect` fires the same commit path the picker's own change handler uses. */
export function createRecentColorsRow(
	doc: Document,
	t: Translator,
	onSelect: (hex: string) => void,
): RecentColorsRowHandle {
	registerPptxWebControls();
	const el = doc.createElement('pptx-ui-color-grid') as PptxUiColorGridElement;
	el.className = 'pptxv-recent-colors';
	el.dataset.testid = 'pptx-color-recent';
	el.addEventListener('office-color-pick', (event) =>
		onSelect((event as OfficeColorPickEvent).detail.color),
	);

	const show = (colors: readonly string[]): void => {
		const view = recentColorsGridView(colors, t);
		el.palette = view.palette;
		el.standardColors = view.standardColors;
		el.recentColors = view.recentColors;
		el.recentHeading = view.recentHeading;
		el.recentLabel = view.recentLabel;
		el.label = view.label;
		el.hidden = colors.length === 0;
	};
	show([]);

	return {
		el,
		setColors(colors) {
			show(colors ?? []);
		},
		setDisabled(disabled) {
			el.disabled = disabled;
		},
	};
}
