import { buildQuickStylesGallery, quickStyleShapeStyle } from 'ooxml-ui/pptx';
import type { PptxUiRibbonGalleryElement, RibbonGalleryPickEvent } from 'ooxml-ui/pptx';
import type { ShapeStyle } from 'pptx-viewer-core';
import React, { useEffect, useLayoutEffect, useRef } from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { LBL } from './FillStrokeHelpers';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

export interface QuickStylesGalleryProps {
	onUpdateShapeStyle: (updates: Partial<ShapeStyle>) => void;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

const QUICK_STYLES = buildQuickStylesGallery();

/**
 * The inspector's Quick Styles: the shared gallery (`pptx-ui-ribbon-gallery`, a compact panel)
 * over the shape style presets. A pick hands the preset's style to `onUpdateShapeStyle`, which
 * merges it over the shape's own.
 */
export function QuickStylesGallery({
	onUpdateShapeStyle,
}: QuickStylesGalleryProps): React.ReactElement {
	const { t } = useTranslation();
	const ref = useRef<PptxUiRibbonGalleryElement>(null);
	useLayoutEffect(() => {
		const gallery = ref.current;
		if (gallery) {
			gallery.descriptor = QUICK_STYLES;
			gallery.translateLabel = t;
		}
	}, [t]);
	useEffect(() => {
		const gallery = ref.current;
		if (!gallery) {
			return;
		}
		const pick = (event: Event) => {
			const style = quickStyleShapeStyle((event as RibbonGalleryPickEvent).detail.itemId);
			if (style) {
				onUpdateShapeStyle(style);
			}
		};
		gallery.addEventListener('gallery-pick', pick);
		return () => gallery.removeEventListener('gallery-pick', pick);
	}, [onUpdateShapeStyle]);

	return (
		<div className='flex flex-col gap-1'>
			<span className={LBL}>{t('pptx.shape.quickStyles')}</span>
			<pptx-ui-ribbon-gallery ref={ref} mode='panel' compact />
		</div>
	);
}
