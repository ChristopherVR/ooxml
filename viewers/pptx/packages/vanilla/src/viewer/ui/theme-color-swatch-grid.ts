import type {
	OfficeColorPickEvent,
	PptxUiColorGridElement,
	ThemeColorPickerCommit,
} from 'ooxml-ui/pptx';
import {
	buildThemeColorSwatchGrid,
	registerPptxWebControls,
	themeColorGridCommit,
	themeColorGridView,
} from 'ooxml-ui/pptx';
import type { PptxThemeColorRef } from 'pptx-viewer-core';

import type { Translator } from '../i18n';

/**
 * PowerPoint's "Theme Colors" grid on the shared `pptx-ui-color-grid`: ten
 * columns (Background 1, Text 1, Background 2, Text 2, Accent 1..6), each a
 * base swatch over five luminance variants, built from the loaded deck's real
 * theme colours rather than a hard-coded Office palette.
 *
 * Mirrors `recent-colors-row.ts`'s handle contract: hidden while there is
 * nothing to show (no theme loaded yet), redrawn on `setThemeColorMap`, and
 * `onPick` fires the caller's own commit path with BOTH the resolved hex and
 * the ref to store. The callers keep their own standard and recent rows, so
 * the grid shows Theme Colors only.
 */
export interface ThemeColorSwatchGridHandle {
	/** The `pptx-ui-color-grid` itself. */
	el: HTMLElement;
	/** Replace the deck's theme colour map (scheme key -> hex); hides the grid when `undefined`. */
	setThemeColorMap(themeColorMap: Record<string, string> | undefined): void;
	/** Highlight the swatch matching the element's current ref/hex, if any. */
	setSelected(ref: PptxThemeColorRef | undefined, hex: string | undefined): void;
	setDisabled(disabled: boolean): void;
}

export function createThemeColorSwatchGrid(
	doc: Document,
	t: Translator,
	onPick: (commit: ThemeColorPickerCommit) => void,
): ThemeColorSwatchGridHandle {
	registerPptxWebControls();
	const el = doc.createElement('pptx-ui-color-grid') as PptxUiColorGridElement;
	const heading = t('pptx.colorPicker.themeColors');
	el.className = 'pptxv-theme-swatch-grid';
	el.standardColors = [];
	el.themeHeading = heading;
	el.label = heading;

	let selectedRef: PptxThemeColorRef | undefined;
	let selectedHex: string | undefined;
	let columns = buildThemeColorSwatchGrid(undefined, t);

	const sync = (): void => {
		const view = themeColorGridView(columns, selectedRef, selectedHex);
		el.palette = view.palette;
		el.selectedTheme = view.selectedTheme;
		el.hidden = columns.length === 0;
	};

	el.addEventListener('office-color-pick', (event) => {
		const commit = themeColorGridCommit(columns, (event as OfficeColorPickEvent).detail);
		if (commit) {
			onPick(commit);
		}
	});
	sync();

	return {
		el,
		setThemeColorMap(themeColorMap) {
			columns = buildThemeColorSwatchGrid(themeColorMap, t);
			sync();
		},
		setSelected(ref, hex) {
			selectedRef = ref;
			selectedHex = hex;
			sync();
		},
		setDisabled(disabled) {
			el.disabled = disabled;
		},
	};
}
