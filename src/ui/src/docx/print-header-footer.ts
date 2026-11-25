import {
	isRunHiddenForReview,
	type DocumentModel,
	type HeaderFooterContent,
	type ReviewDisplayMode,
} from 'ooxml-core/docx';
import {
	cssFontStack,
	createPageStories,
	pageNumberValues,
	headerFooterSlotForPage,
	type LayoutPageBox,
	type TextMeasurer,
} from 'ooxml-core/docx/layout';
import { printPicture, renderStoryBlocks, type PictureUrl } from './print-layout';
import { createCanvasMeasurer } from './canvas-measurer';
import { pictureWatermarkElement, type WatermarkPage } from './print-picture-watermark';

export { pageNumbers, headerFooterForPage, fieldDisplayText } from 'ooxml-core/docx/layout';

export function headerFooterElement(kind: 'header' | 'footer'): HTMLElement {
	const element = document.createElement('div');
	element.className = `dve-print-${kind}`;
	element.setAttribute('aria-label', kind === 'header' ? 'Header' : 'Footer');
	return element;
}

/**
 * The header's watermark as a layer behind the page text, or null: text centred and rotated as
 * Word sets it, or a picture centred on the margins.
 */
export function watermarkElement(
	content: HeaderFooterContent,
	page: WatermarkPage,
	reviewDisplayMode: ReviewDisplayMode = 'all',
	pictureUrl?: PictureUrl,
): HTMLElement | null {
	const image = content.blocks
		.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.find(
			(run) =>
				!isRunHiddenForReview(run, reviewDisplayMode) &&
				(run.image?.watermark || run.image?.pictureWatermark),
		)?.image;
	const picture = image?.pictureWatermark;
	if (picture)
		return pictureWatermarkElement(
			picture,
			page,
			pictureUrl?.(picture.partName, picture.contentType),
		);
	const spec = image?.watermark;
	if (!spec?.text) return null;
	const chars = Math.max(1, [...spec.text].length);
	const span = Math.min(page.widthPx * (spec.layout === 'diagonal' ? 0.78 : 0.7), 560);
	const layer = document.createElement('div');
	layer.className = 'dve-print-watermark';
	layer.setAttribute('aria-hidden', 'true');
	layer.textContent = spec.text;
	Object.assign(layer.style, {
		position: 'absolute',
		left: '0',
		top: '0',
		width: `${page.widthPx}px`,
		height: `${page.heightPx}px`,
		display: 'flex',
		alignItems: 'center',
		justifyContent: 'center',
		pointerEvents: 'none',
		userSelect: 'none',
		WebkitUserSelect: 'none',
		whiteSpace: 'nowrap',
		overflow: 'hidden',
		fontSize: `${Math.max(12, Math.min(220, span / (chars * 0.6)))}px`,
		fontWeight: '700',
		color: /^#[0-9a-f]{6}$/i.test(spec.color) ? spec.color : '#c0c0c0',
		opacity: spec.semitransparent ? '0.5' : '1',
		transform: spec.layout === 'diagonal' ? 'rotate(-45deg)' : 'none',
		zIndex: '0',
		...(spec.fontFamily ? { fontFamily: cssFontStack(spec.fontFamily) } : {}),
	});
	return layer;
}

/** Creates a decorator for both eagerly rendered and virtualized pages. */
export function createPageDecorator(
	model: DocumentModel,
	pages: LayoutPageBox[],
	now: Date,
	pictureUrl: PictureUrl | undefined,
	reviewDisplayMode: ReviewDisplayMode,
	measurer: TextMeasurer,
) {
	const numbers = pageNumberValues(model, pages);
	const stories = createPageStories(model, pages, measurer, { reviewDisplayMode }, undefined, now);
	return (sheet: HTMLElement, index: number) => {
		const page = pages[index];
		if (!page) return;
		const section = model.sections?.[page.sectionIndex];
		Object.assign(sheet.dataset, {
			sectionIndex: String(page.sectionIndex),
			headerFooterSlot: headerFooterSlotForPage(model, page, numbers[index] ?? index + 1),
			marginTop: String(page.marginTopPx),
			marginBottom: String(page.marginBottomPx),
			marginLeft: String(page.marginLeftPx),
			marginRight: String(page.marginRightPx),
			headerDistance: String((section?.headerDistanceTwips ?? 720) / 15),
			footerDistance: String((section?.footerDistanceTwips ?? 720) / 15),
		});
		const media = { model, pictureUrl };
		for (const story of stories(index)) {
			const element = headerFooterElement(story.kind);
			element.style.left = `${story.x}px`;
			element.style.width = `${story.widthPx}px`;
			element.style.height = `${story.heightPx}px`;
			element.style[story.kind === 'header' ? 'top' : 'bottom'] = `${story.distancePx}px`;
			element.append(...renderStoryBlocks(story.blocks, media, story.widthPx));
			sheet.append(element);
			for (const float of story.floats) {
				const picture = printPicture(float, media);
				picture.classList.add('dve-print-float');
				if (float.behindText) picture.classList.add('dve-print-float-behind');
				picture.style.left = `${float.xPx}px`;
				picture.style.top = `${float.yPx}px`;
				sheet.append(picture);
			}
			if (story.kind === 'header') {
				const mark = watermarkElement(story.content, page, reviewDisplayMode, pictureUrl);
				if (mark) sheet.prepend(mark);
			}
		}
	};
}

export function decoratePages(
	model: DocumentModel,
	pages: LayoutPageBox[],
	sheets: HTMLElement[],
	now = new Date(),
	pictureUrl?: PictureUrl,
	reviewDisplayMode: ReviewDisplayMode = 'all',
) {
	sheets.forEach(
		createPageDecorator(model, pages, now, pictureUrl, reviewDisplayMode, createCanvasMeasurer()),
	);
}
