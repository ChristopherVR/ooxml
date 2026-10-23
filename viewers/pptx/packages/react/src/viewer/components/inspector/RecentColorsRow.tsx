import type { OfficeColorPickEvent, PptxUiColorGridElement } from 'ooxml-ui/pptx';
import { recentColorsGridView } from 'ooxml-ui/pptx';
import React, { useEffect, useLayoutEffect, useRef } from 'react';
import { useTranslation } from 'react-i18next';

import { useRecentColors } from './RecentColorsContext';

// ---------------------------------------------------------------------------
// RecentColorsRow
// ---------------------------------------------------------------------------

/**
 * The "Recent colours" row under every colour picker in the inspector, drawn
 * by the shared `pptx-ui-color-grid` with only its Recent Colors section
 * ({@link recentColorsGridView}) and sourced from {@link useRecentColors}
 * (context), so every caller renders the SAME list rather than re-deriving it.
 *
 * Renders nothing while the list is empty, so a caller can render it
 * unconditionally. Clicking a swatch both calls `onCommit` (the picker's own
 * apply path) AND pushes the colour back to the front of the shared list
 * itself, so every caller gets the row's full contract for free without also
 * having to remember to push.
 */
export function RecentColorsRow({
	disabled,
	onCommit,
}: {
	disabled?: boolean;
	/** Applies the picked colour through the picker's own commit path. */
	onCommit: (color: string) => void;
}): React.ReactElement | null {
	const { t } = useTranslation();
	const { recentColors, pushColor } = useRecentColors();
	const shown = recentColors.length > 0;
	const ref = useRef<PptxUiColorGridElement>(null);
	useLayoutEffect(() => {
		const grid = ref.current;
		if (!grid) {
			return;
		}
		const view = recentColorsGridView(recentColors, (key, params) => t(key, params));
		grid.palette = view.palette;
		grid.standardColors = view.standardColors;
		grid.recentColors = view.recentColors;
		grid.recentHeading = view.recentHeading;
		grid.recentLabel = view.recentLabel;
		grid.label = view.label;
		grid.disabled = Boolean(disabled);
	}, [recentColors, disabled, t]);
	// The grid only exists while there is a recent colour, so the listener attaches once shown.
	useEffect(() => {
		const grid = ref.current;
		if (!shown || !grid) {
			return;
		}
		const pick = (event: Event) => {
			const { color } = (event as OfficeColorPickEvent).detail;
			onCommit(color);
			pushColor(color);
		};
		grid.addEventListener('office-color-pick', pick);
		return () => grid.removeEventListener('office-color-pick', pick);
	}, [shown, onCommit, pushColor]);

	if (!shown) {
		return null;
	}
	return <pptx-ui-color-grid ref={ref} data-testid='pptx-color-recent' className='mt-1' />;
}
