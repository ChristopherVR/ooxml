import type { PageBorders } from 'ooxml-core/docx';
import { cssBorderSide } from 'ooxml-core/docx/ui';

/** The page geometry the border box is measured from (pixels at 96 DPI). */
export interface BorderPage {
	widthPx: number;
	heightPx: number;
	marginTopPx: number;
	marginRightPx: number;
	marginBottomPx: number;
	marginLeftPx: number;
	pageInSection: number;
}

const POINT = 4 / 3;

/**
 * The box a section's page borders are drawn inside on one page, or null when they do not show on
 * it (`w:display`). Each side sits `w:space` points from the page edge (default) or from the text
 * margins (`offsetFrom="text"`). DOM-free: the PDF export draws the same box.
 */
export function pageBorderRect(
	page: BorderPage,
	borders: PageBorders,
): { leftPx: number; topPx: number; widthPx: number; heightPx: number } | null {
	if (borders.display === 'firstPage' && page.pageInSection !== 0) return null;
	if (borders.display === 'notFirstPage' && page.pageInSection === 0) return null;
	const fromText = borders.offsetFrom === 'text';
	// From the page edge the space reaches the line's outer edge; from the text, its inner edge
	// (measured in Word's PDF: a 1.5 pt line 24 pt from 1 in margins spans 46.5 to 48 pt).
	const space = (side: 'top' | 'left' | 'bottom' | 'right') =>
		((borders[side]?.spacePoints ?? 24) +
			(fromText ? (borders[side]?.sizeEighthPoints ?? 4) / 8 : 0)) *
		POINT;
	const left = Math.max(0, fromText ? page.marginLeftPx - space('left') : space('left'));
	const top = Math.max(0, fromText ? page.marginTopPx - space('top') : space('top'));
	const right = Math.max(0, fromText ? page.marginRightPx - space('right') : space('right'));
	const bottom = Math.max(0, fromText ? page.marginBottomPx - space('bottom') : space('bottom'));
	return {
		leftPx: left,
		topPx: top,
		widthPx: Math.max(0, page.widthPx - left - right),
		heightPx: Math.max(0, page.heightPx - top - bottom),
	};
}

/**
 * A positioned box drawing a section's page borders on one printed sheet, or null when the borders
 * do not show on this page (`w:display`) or no side has a line. Each side sits `w:space` points from
 * the page edge (default) or from the text margins (`offsetFrom="text"`). Art borders are drawn as
 * plain lines, as the document model reports. It never takes pointer events.
 */
export function pageBorderBox(page: BorderPage, borders: PageBorders): HTMLElement | null {
	const css = {
		top: cssBorderSide(borders.top),
		left: cssBorderSide(borders.left),
		bottom: cssBorderSide(borders.bottom),
		right: cssBorderSide(borders.right),
	};
	const rect = pageBorderRect(page, borders);
	if (!rect || !Object.values(css).some(Boolean)) return null;
	const box = document.createElement('div');
	box.className = 'dve-print-page-border';
	box.setAttribute('aria-hidden', 'true');
	Object.assign(box.style, {
		position: 'absolute',
		boxSizing: 'border-box',
		pointerEvents: 'none',
		left: `${rect.leftPx}px`,
		top: `${rect.topPx}px`,
		width: `${rect.widthPx}px`,
		height: `${rect.heightPx}px`,
		zIndex: borders.zOrder === 'back' ? '0' : '2',
	});
	for (const [side, value] of Object.entries(css))
		if (value) box.style.setProperty(`border-${side}`, value);
	return box;
}
