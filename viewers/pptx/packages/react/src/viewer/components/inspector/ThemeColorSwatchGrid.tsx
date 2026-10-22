import type {
	OfficeColorPickEvent,
	PptxUiColorGridElement,
	ThemeColorPickerCommit,
} from 'ooxml-ui/pptx';
import { buildThemeColorSwatchGrid, themeColorGridCommit, themeColorGridView } from 'ooxml-ui/pptx';
import type { PptxThemeColorRef } from 'pptx-viewer-core';
import React, { useEffect, useLayoutEffect, useMemo, useRef } from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { useThemeColorMap } from './ThemeColorMapContext';

export interface ThemeColorSwatchGridProps {
	disabled?: boolean;
	/** The element's current theme ref, if any (used to highlight the matching swatch). */
	selectedRef?: PptxThemeColorRef;
	/** The element's current resolved hex, used to highlight a swatch when no ref is stored. */
	selectedHex?: string;
	/** Called when the user clicks a swatch: both the resolved hex and the ref to store. */
	onPick: (commit: ThemeColorPickerCommit) => void;
}

/**
 * PowerPoint's "Theme Colors" grid on the shared `pptx-ui-color-grid`: ten columns (Background 1,
 * Text 1, Background 2, Text 2, Accent 1..6), each a base swatch over five luminance variants,
 * built from the loaded deck's theme ({@link useThemeColorMap}) rather than a hard-coded Office
 * palette.
 *
 * Renders nothing when no deck theme is loaded yet, so callers can render this unconditionally
 * alongside their existing hex and recent-colour controls.
 */
export function ThemeColorSwatchGrid({
	disabled,
	selectedRef,
	selectedHex,
	onPick,
}: ThemeColorSwatchGridProps): React.ReactElement | null {
	const { t } = useTranslation();
	const themeColorMap = useThemeColorMap();
	const columns = useMemo(
		() => buildThemeColorSwatchGrid(themeColorMap, (key, params) => t(key, params)),
		[themeColorMap, t],
	);
	const ref = useRef<PptxUiColorGridElement>(null);
	useLayoutEffect(() => {
		const grid = ref.current;
		if (!grid) {
			return;
		}
		const view = themeColorGridView(columns, selectedRef, selectedHex);
		const heading = t('pptx.colorPicker.themeColors');
		grid.palette = view.palette;
		grid.selectedTheme = view.selectedTheme;
		grid.standardColors = [];
		grid.themeHeading = heading;
		grid.label = heading;
		grid.disabled = Boolean(disabled);
	}, [columns, selectedRef, selectedHex, disabled, t]);
	useEffect(() => {
		const grid = ref.current;
		if (!grid) {
			return;
		}
		const pick = (event: Event) => {
			const commit = themeColorGridCommit(columns, (event as OfficeColorPickEvent).detail);
			if (commit) {
				onPick(commit);
			}
		};
		grid.addEventListener('office-color-pick', pick);
		return () => grid.removeEventListener('office-color-pick', pick);
	}, [columns, onPick]);

	if (columns.length === 0) {
		return null;
	}
	return <pptx-ui-color-grid ref={ref} className='mt-1' />;
}
