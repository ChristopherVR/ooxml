import type { PictureWatermarkSpec } from 'ooxml-core/docx';
import type { LayoutPageBox } from 'ooxml-core/docx/layout';

/**
 * Washout as a CSS filter: DrawingML's washout (contrast -70%, then brightness +70%) maps every
 * channel x to 0.09x + 0.805, which `contrast(0.053) brightness(1.7)` reproduces.
 */
export const WASHOUT_FILTER = 'contrast(0.053) brightness(1.7)';

/** The page box a watermark is drawn on; without margins a picture centres on the page. */
export type WatermarkPage = Pick<LayoutPageBox, 'widthPx' | 'heightPx'> &
	Partial<Pick<LayoutPageBox, 'marginLeftPx' | 'marginRightPx' | 'marginTopPx' | 'marginBottomPx'>>;

/** A header's picture watermark, centred on the page's margins behind the text. */
export function pictureWatermarkElement(
	spec: PictureWatermarkSpec,
	page: WatermarkPage,
	url: string | undefined,
): HTMLElement {
	const widthPx = (spec.widthPt * 4) / 3;
	const heightPx = (spec.heightPt * 4) / 3;
	const left = page.marginLeftPx ?? 0;
	const top = page.marginTopPx ?? 0;
	const areaWidth = page.widthPx - left - (page.marginRightPx ?? 0);
	const areaHeight = page.heightPx - top - (page.marginBottomPx ?? 0);
	const layer = document.createElement(url ? 'img' : 'div');
	if (url) (layer as HTMLImageElement).src = url;
	else layer.classList.add('dve-print-picture-missing');
	layer.className = `${layer.className} dve-print-watermark dve-print-watermark-picture`.trim();
	layer.setAttribute('aria-hidden', 'true');
	if (url) (layer as HTMLImageElement).alt = '';
	Object.assign(layer.style, {
		position: 'absolute',
		left: `${left + (areaWidth - widthPx) / 2}px`,
		top: `${top + (areaHeight - heightPx) / 2}px`,
		width: `${widthPx}px`,
		height: `${heightPx}px`,
		pointerEvents: 'none',
		userSelect: 'none',
		WebkitUserSelect: 'none',
		zIndex: '0',
		...(spec.washout ? { filter: WASHOUT_FILTER } : {}),
	});
	return layer;
}
