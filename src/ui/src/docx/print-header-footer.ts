import { pictureNeedsSvg, pictureSvg } from '../picture-view';
import {
	drawingShapeSvg,
	fieldName,
	isRunHiddenForReview,
	refreshComputedFields,
	type Block,
	type DocumentModel,
	type HeaderFooterContent,
	type Paragraph,
	type TextRun,
	type ReviewDisplayMode,
} from 'ooxml-core/docx';
import {
	cssFontStack,
	floatPosition,
	paragraphFloats,
	pageNumbers,
	pageNumberValues,
	sectionPageCounts,
	headerFooterForPage,
	headerFooterSlotForPage,
	fieldDisplayText,
	pageStyleRefs,
	type PageFieldContext,
	type LayoutPageBox,
} from 'ooxml-core/docx/layout';
import type { PictureUrl } from './print-layout';
import { pictureElement } from './print-picture';
import { pictureWatermarkElement, type WatermarkPage } from './print-picture-watermark';

/** Page facts a header or footer field can show. */
export interface PageFieldValues extends PageFieldContext {
	reviewDisplayMode?: ReviewDisplayMode;
	/** Resolves header/footer pictures to displayable URLs. */
	pictureUrl?: PictureUrl;
	/** The document: its theme colours shapes, its styles a text box's paragraphs. */
	model?: DocumentModel;
}

export { pageNumbers, headerFooterForPage, fieldDisplayText } from 'ooxml-core/docx/layout';

/** An inline header/footer shape, drawn as Print Layout draws one in the body. */
function inlineShape(image: NonNullable<TextRun['image']>, values: PageFieldValues) {
	const svg = drawingShapeSvg(image, values.model?.theme);
	if (!svg) return undefined;
	const element = pictureElement(
		{
			...image,
			chart: { svg, label: 'Shape', shape: true },
			...(image.textBoxContent ? { textBox: image.textBoxContent } : {}),
		},
		undefined,
		values.model,
	);
	Object.assign(element.style, { position: 'relative', display: 'inline-block' });
	element.style.verticalAlign = 'bottom';
	return element;
}

/** An inline header/footer picture, or a placeholder box when its bytes are unavailable. */
function inlinePicture(image: NonNullable<TextRun['image']>, values: PageFieldValues): HTMLElement {
	const shape = inlineShape(image, values);
	if (shape) return shape;
	const url = image.partName ? values.pictureUrl?.(image.partName, image.contentType) : undefined;
	const element = document.createElement(
		url && !pictureNeedsSvg(image.pictureFill) ? 'img' : 'span',
	);
	if (url && image.pictureFill && pictureNeedsSvg(image.pictureFill)) {
		element.append(pictureSvg(document, image.pictureFill, url, image.widthPx, image.heightPx));
		element.setAttribute('aria-label', image.altText ?? '');
	} else if (url) {
		(element as HTMLImageElement).src = url;
		(element as HTMLImageElement).alt = image.altText ?? '';
	} else element.className = 'dve-print-picture-missing';
	element.style.display = 'inline-block';
	element.style.verticalAlign = 'bottom';
	element.style.width = `${image.widthPx}px`;
	element.style.height = `${image.heightPx}px`;
	return element;
}

function runElement(run: TextRun, values: PageFieldValues): HTMLElement {
	if (run.image) return inlinePicture(run.image, values);
	const span = document.createElement('span');
	span.textContent = fieldDisplayText(run, values);
	if (run.field) span.dataset.field = fieldName(run.field.instr);
	if (run.bold) span.style.fontWeight = '700';
	if (run.italic) span.style.fontStyle = 'italic';
	if (run.underline) span.style.textDecoration = 'underline';
	if (run.fontSize) span.style.fontSize = `${run.fontSize}pt`;
	if (run.fontFamily) span.style.fontFamily = cssFontStack(run.fontFamily);
	if (run.color && /^#[0-9a-f]{6}$/i.test(run.color)) span.style.color = run.color;
	return span;
}

function paragraphElement(paragraph: Paragraph, values: PageFieldValues): HTMLElement {
	const element = document.createElement('p');
	if (paragraph.align) element.style.textAlign = paragraph.align;
	if (paragraph.direction) element.dir = paragraph.direction;
	// Floating pictures are placed on the sheet by `decoratePages`, not in the text flow.
	for (const run of paragraph.runs)
		if (
			!isRunHiddenForReview(run, values.reviewDisplayMode ?? 'all') &&
			!run.break &&
			!run.image?.anchored &&
			!run.image?.watermark &&
			!run.image?.pictureWatermark
		)
			element.append(runElement(run, values));
	return element;
}

function blockElement(block: Block, values: PageFieldValues): HTMLElement {
	if (block.type === 'paragraph') return paragraphElement(block, values);
	const table = document.createElement('table');
	for (const row of block.rows) {
		const tr = document.createElement('tr');
		for (const cell of row) {
			const td = document.createElement('td');
			if (cell.gridSpan && cell.gridSpan > 1) td.colSpan = cell.gridSpan;
			for (const paragraph of cell.paragraphs) td.append(paragraphElement(paragraph, values));
			tr.append(td);
		}
		table.append(tr);
	}
	return table;
}

/** Renders read-only header/footer content with page fields filled in for one page. */
export function renderHeaderFooter(
	content: HeaderFooterContent,
	values: PageFieldValues,
	kind: 'header' | 'footer',
): HTMLElement {
	const element = document.createElement('div');
	element.className = `dve-print-${kind}`;
	element.setAttribute('aria-label', kind === 'header' ? 'Header' : 'Footer');
	for (const block of content.blocks) element.append(blockElement(block, values));
	return element;
}

/** Adds each page's header and footer, positioned at the section's header/footer distances. */
/** Header/footer floating pictures, positioned on the sheet from their anchor frames. */
function floatingPictures(
	content: HeaderFooterContent,
	page: LayoutPageBox,
	anchorTopPx: number,
	values: PageFieldValues,
): HTMLElement[] {
	const paragraphs = content.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
	);
	const column = { xPx: 0, widthPx: page.widthPx - page.marginLeftPx - page.marginRightPx };
	return paragraphs.flatMap((paragraph) =>
		paragraphFloats(
			{
				...paragraph,
				runs: paragraph.runs.filter(
					(run) => !isRunHiddenForReview(run, values.reviewDisplayMode ?? 'all'),
				),
			},
			values.model?.theme,
		).map((float) => {
			const { xPx, yPx } = floatPosition(float, page, column, { topPx: anchorTopPx, heightPx: 0 });
			// The body's floating pictures and shapes are drawn by the same function.
			const element = pictureElement(float, values.pictureUrl, values.model);
			element.classList.add('dve-print-float');
			if (float.behindText) element.classList.add('dve-print-float-behind');
			element.style.left = `${xPx}px`;
			element.style.top = `${yPx}px`;
			return element;
		}),
	);
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

export function decoratePages(
	model: DocumentModel,
	pages: LayoutPageBox[],
	sheets: HTMLElement[],
	now: Date = new Date(),
	pictureUrl?: PictureUrl,
	reviewDisplayMode: ReviewDisplayMode = 'all',
) {
	// Header and footer DOCPROPERTY results follow the document's properties.
	model = refreshComputedFields(model);
	const numbers = pageNumbers(model, pages);
	const numericValues = pageNumberValues(model, pages);
	const counts = sectionPageCounts(pages);
	// STYLEREF fields show the text of the page they are drawn on, as in Word.
	const styleRefs = pageStyleRefs(model, pages);
	pages.forEach((page, index) => {
		const sheet = sheets[index];
		if (!sheet) return;
		const label = numbers[index] ?? String(index + 1);
		const values: PageFieldValues = {
			reviewDisplayMode,
			model,
			page: label,
			numPages: String(pages.length),
			sectionPages: String(counts.get(page.sectionIndex) ?? 1),
			now,
			...(pictureUrl ? { pictureUrl } : {}),
		};
		const section = model.sections?.[page.sectionIndex];
		const pageNumber = numericValues[index] ?? index + 1;
		// What a double-click on the page's header or footer area needs to edit the right story.
		Object.assign(sheet.dataset, {
			sectionIndex: String(page.sectionIndex),
			headerFooterSlot: headerFooterSlotForPage(model, page, pageNumber),
			marginTop: String(page.marginTopPx),
			marginBottom: String(page.marginBottomPx),
			marginLeft: String(page.marginLeftPx),
			marginRight: String(page.marginRightPx),
			headerDistance: String((section?.headerDistanceTwips ?? 720) / 15),
			footerDistance: String((section?.footerDistanceTwips ?? 720) / 15),
		});
		for (const kind of ['header', 'footer'] as const) {
			const content = headerFooterForPage(model, page, pageNumber, `${kind}s`);
			if (!content) continue;
			const element = renderHeaderFooter(styleRefs.content(content, index), values, kind);
			const distanceTwips =
				kind === 'header' ? section?.headerDistanceTwips : section?.footerDistanceTwips;
			const distancePx = (distanceTwips ?? 720) / 15;
			element.style.left = `${page.marginLeftPx}px`;
			element.style.right = `${page.marginRightPx}px`;
			element.style[kind === 'header' ? 'top' : 'bottom'] = `${distancePx}px`;
			sheet.append(element);
			// Anchors in a footer sit roughly one line above its bottom distance.
			const anchorTop = kind === 'header' ? distancePx : page.heightPx - distancePx - 20;
			sheet.append(...floatingPictures(content, page, anchorTop, values));
			if (kind === 'header') {
				const mark = watermarkElement(content, page, reviewDisplayMode, values.pictureUrl);
				// First in the sheet so the page text paints over it, as a Word watermark sits behind text.
				if (mark) sheet.prepend(mark);
			}
		}
	});
}
