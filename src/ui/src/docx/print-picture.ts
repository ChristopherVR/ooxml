import { pictureNeedsSvg, pictureSvg } from '../picture-view';
import type { DrawingFill, DrawingSvgShadow } from 'ooxml-core/drawingml';
import type { DocumentModel, TextBoxContent } from 'ooxml-core/docx';
import type { TextBoxLayout } from 'ooxml-core/docx/layout';
import { paintWordArtEffects, textBoxContentElement, wordArtWarpElement } from './story-render';
// Print Layout pictures: package media through `PictureUrl`, charts and shapes painted by the core
// (`LayoutObject.chart`) with a shape's text box paragraphs over it (the lines the page engine laid
// out, or the paragraphs flowed by CSS where it laid out none), and a placeholder box for anything
// that cannot be shown.

/** Draws a text box's laid-out lines over its shape, in the colour its runs take where they set none. */
export type TextBoxLayer = (layout: TextBoxLayout, color: string | undefined) => HTMLElement;

/** Resolves a package picture part to a displayable URL; undefined draws a placeholder box. */
export type PictureUrl = (partName: string, contentType: string) => string | undefined;

/**
 * Whether an SVG `href` / `xlink:href` may stay: an in-document fragment or an
 * embedded raster image. Anything else (a script URL, a remote resource) goes.
 */
function isSafeSvgHref(value: string): boolean {
	const trimmed = value.trim();
	return trimmed.startsWith('#') || /^data:image\//iu.test(trimmed);
}

/**
 * Strips the active content from a chart SVG before it is imported into the
 * page: `<script>` and `<foreignObject>` elements, every `on*` event-handler
 * attribute, and `href` / `xlink:href` values other than fragments or
 * `data:image` URLs. Mirrors the rules of `src/core/docx/ui/chart-spec.ts`.
 */
export function sanitizeChartSvg(svg: Element): void {
	const elements = [svg, ...Array.from(svg.getElementsByTagName('*'))];
	for (const element of elements) {
		const name = element.localName.toLowerCase();
		if (name === 'script' || name === 'foreignobject') {
			element.remove();
			continue;
		}
		for (const attr of Array.from(element.attributes)) {
			const unsafeHref = attr.localName === 'href' && !isSafeSvgHref(attr.value);
			if (/^on/iu.test(attr.name) || unsafeHref) {
				element.removeAttribute(attr.name);
			}
		}
	}
}

/** A chart painted by the core (`LayoutObject.chart`), or its labelled frame. */
function chartElement(chart: { svg?: string; label: string; shape?: true }): HTMLElement {
	const element = document.createElement('div');
	element.dataset.docxChart = chart.svg ? '1' : 'frame';
	const svg = chart.svg
		? new DOMParser().parseFromString(chart.svg, 'image/svg+xml').documentElement
		: undefined;
	if (svg && svg.localName === 'svg') {
		sanitizeChartSvg(svg);
		element.classList.add('dve-print-chart');
		// A shape's centred outline straddles its box, and it has no backdrop of
		// its own: the chart frame's clip halved the outline and its white fill
		// showed in the corners around a non-rectangular shape.
		if (chart.shape) element.classList.add('dve-print-shape');
		element.title = chart.label;
		element.append(document.importNode(svg, true));
	} else {
		element.classList.add('dve-print-picture-missing', 'dve-print-chart-frame');
		element.textContent = chart.label;
	}
	return element;
}

function pictureShadowCss(shadow: DrawingSvgShadow): string {
	const alpha = Math.round(Math.max(0, Math.min(1, shadow.opacity)) * 255)
		.toString(16)
		.padStart(2, '0');
	const color = shadow.color.startsWith('#') ? shadow.color : `#${shadow.color}`;
	return `drop-shadow(${shadow.dx}px ${shadow.dy}px ${shadow.blur / 2}px ${color}${alpha})`;
}

export function pictureElement(
	object: {
		altText?: string;
		transform?: import('ooxml-core/docx').DrawingTransform;
		pictureFill?: Extract<DrawingFill, { kind: 'picture' }>;
		shadow?: DrawingSvgShadow;
		partName: string;
		contentType: string;
		widthPx: number;
		heightPx: number;
		chart?: { svg?: string; label: string; shape?: true; textColor?: string };
		textBox?: TextBoxContent;
		textBoxLayout?: TextBoxLayout;
		noteSeparator?: 'separator' | 'continuationSeparator';
	},
	pictureUrl: PictureUrl | undefined,
	/** The document, whose styles a text box's paragraphs are drawn with. */
	model?: DocumentModel,
	/** Draws `textBoxLayout`; without it, or without a layout, the paragraphs are flowed by CSS. */
	textBoxLayer?: TextBoxLayer,
): HTMLElement {
	if (object.noteSeparator) {
		const rule = document.createElement('span');
		rule.className = 'dve-print-note-rule';
		Object.assign(rule.style, {
			position: 'absolute',
			width: `${object.widthPx}px`,
			height: '1px',
			borderTop: '1px solid currentColor',
		});
		return rule;
	}
	if (object.chart) {
		const element = chartElement(object.chart);
		if (object.altText) element.setAttribute('aria-label', object.altText);
		element.classList.add('dve-print-picture');
		element.style.position = 'absolute';
		element.style.width = `${object.widthPx}px`;
		element.style.height = `${object.heightPx}px`;
		if (object.chart.svg && object.chart.shape && object.textBox) {
			const warp = wordArtWarpElement(object.textBox, object.widthPx, object.heightPx);
			if (warp) {
				element.append(warp);
			} else if (object.textBoxLayout && textBoxLayer) {
				const layer = textBoxLayer(object.textBoxLayout, object.chart.textColor);
				paintWordArtEffects(layer, object.textBox);
				element.append(layer);
			} else {
				element.append(
					textBoxContentElement(object.textBox, model, object.chart.textColor, pictureUrl),
				);
			}
		}
		return element;
	}
	const url = object.partName ? pictureUrl?.(object.partName, object.contentType) : undefined;
	if (url && object.pictureFill && pictureNeedsSvg(object.pictureFill)) {
		const element = document.createElement('span');
		element.className = 'dve-print-picture';
		element.style.position = 'absolute';
		element.style.width = `${object.widthPx}px`;
		element.style.height = `${object.heightPx}px`;
		if (object.altText) {
			element.setAttribute('role', 'img');
			element.setAttribute('aria-label', object.altText);
		}
		element.style.transform = pictureTransformCss(object.transform);
		if (object.shadow) element.style.filter = pictureShadowCss(object.shadow);
		element.append(pictureSvg(document, object.pictureFill, url, object.widthPx, object.heightPx));
		return element;
	}
	const element = document.createElement(url ? 'img' : 'div');
	if (url) {
		(element as HTMLImageElement).src = url;
		(element as HTMLImageElement).alt = object.altText ?? '';
	} else element.classList.add('dve-print-picture-missing');
	element.classList.add('dve-print-picture');
	element.style.position = 'absolute';
	element.style.width = `${object.widthPx}px`;
	element.style.height = `${object.heightPx}px`;
	element.style.transform = pictureTransformCss(object.transform);
	if (object.shadow) element.style.filter = pictureShadowCss(object.shadow);
	return element;
}
import { pictureTransformCss } from './picture-transform';
