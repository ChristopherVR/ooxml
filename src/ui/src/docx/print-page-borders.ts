import type { PageBorders } from 'ooxml-core/docx';
import { pageBorderRect, type BorderPage } from 'ooxml-core/docx/pdf-export';
import { cssBorderSide } from 'ooxml-core/docx/ui';

export { pageBorderRect, type BorderPage } from 'ooxml-core/docx/pdf-export';

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
