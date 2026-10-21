import {
	buildTextWarpGallery,
	TEXT_WARP_PRESETS,
	textWarpStyle,
	warpPreviewPath,
} from 'ooxml-ui/pptx';
import type { PptxUiRibbonGalleryElement, RibbonGalleryPickEvent } from 'ooxml-ui/pptx';
import type { TextStyle } from 'pptx-viewer-core';
import React, { useEffect, useLayoutEffect, useMemo, useRef } from 'react';
import { useTranslation } from 'react-i18next';

// The warp preset catalogue + preview-path generator now live in
// `ooxml-ui/pptx` (shared by every binding). Re-export them here to preserve
// this module's historical public symbol surface.
export { TEXT_WARP_PRESETS, warpPreviewPath };

// ==========================================================================
// Component
// ==========================================================================

interface TextWarpGalleryProps {
	ts: TextStyle | undefined;
	onUpdateTextStyle: (updates: Partial<TextStyle>) => void;
}

/**
 * The inspector's Text Warp section: a collapsible header over the shared Text Warp gallery
 * (`pptx-ui-ribbon-gallery`, a compact panel), which stays in the DOM while collapsed.
 */
export function TextWarpGallery({
	ts,
	onUpdateTextStyle,
}: TextWarpGalleryProps): React.ReactElement {
	const { t } = useTranslation();
	const [expanded, setExpanded] = React.useState(false);
	const currentPreset = ts?.textWarpPreset || 'textNoShape';
	const descriptor = useMemo(() => buildTextWarpGallery(ts?.textWarpPreset), [ts?.textWarpPreset]);
	const ref = useRef<PptxUiRibbonGalleryElement>(null);
	useLayoutEffect(() => {
		const gallery = ref.current;
		if (gallery) {
			gallery.descriptor = descriptor;
			gallery.translateLabel = t;
		}
	}, [descriptor, t]);
	useEffect(() => {
		const gallery = ref.current;
		if (!gallery) {
			return;
		}
		const pick = (event: Event) => {
			const style = textWarpStyle((event as RibbonGalleryPickEvent).detail.itemId);
			if (style) {
				onUpdateTextStyle(style);
			}
		};
		gallery.addEventListener('gallery-pick', pick);
		return () => gallery.removeEventListener('gallery-pick', pick);
	}, [onUpdateTextStyle]);

	return (
		<div className='mt-2 rounded border border-border bg-card p-2 space-y-2'>
			<button
				type='button'
				className='flex w-full items-center justify-between text-[11px] uppercase tracking-wide text-muted-foreground max-md:min-h-[44px]!'
				onClick={() => setExpanded(!expanded)}
			>
				<span>{t('pptx.textWarp.title')}</span>
				<span className='text-muted-foreground'>{expanded ? '−' : '+'}</span>
			</button>
			{!expanded && (
				<div className='text-[11px] text-muted-foreground'>
					{(() => {
						const current = TEXT_WARP_PRESETS.find((p) => p.value === currentPreset);
						return current ? t(current.labelKey) : currentPreset;
					})()}
				</div>
			)}
			{/* A plain wrapper hides it: React 18 writes hidden={false} on a custom element as "false". */}
			<div hidden={!expanded}>
				<pptx-ui-ribbon-gallery ref={ref} mode='panel' compact />
			</div>
		</div>
	);
}
