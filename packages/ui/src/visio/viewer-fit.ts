import type { VisioPage } from 'ooxml-core/visio';

/**
 * Zoom that fits a page into the drawing window, as Visio's Fit to Window (whole page, never
 * enlarged past 100%) or Page Width (fills the width and scrolls vertically).
 */
export function fitZoom(viewport: HTMLElement, page: VisioPage, mode: 'page' | 'width'): number {
	const style = viewport.ownerDocument.defaultView?.getComputedStyle(viewport);
	const padding = (value: string | undefined) => Number.parseFloat(value ?? '') || 0;
	const width = Math.max(
		1,
		viewport.clientWidth - padding(style?.paddingLeft) - padding(style?.paddingRight),
	);
	const height = Math.max(
		1,
		viewport.clientHeight - padding(style?.paddingTop) - padding(style?.paddingBottom),
	);
	if (mode === 'page') return Math.min(1, width / (page.width * 96), height / (page.height * 96));
	// Page Width scrolls vertically; reserve that scrollbar so no horizontal one appears.
	const scrollbar = viewport.offsetWidth - viewport.clientWidth;
	const fitted = width / (page.width * 96);
	const reserve = page.height * 96 * fitted > height && scrollbar <= 0 ? 17 : 0;
	return Math.max(1, width - reserve - 1) / (page.width * 96);
}
