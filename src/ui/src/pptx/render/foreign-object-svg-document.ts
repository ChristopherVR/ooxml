import { settleThreeViews, snapshotThreeViewsIntoClone } from '../three-view/export-snapshot';
import { withAuthoredStrokeWidths } from './device-pixel-stroke';
import { prepareExportClone } from './export-clone';
import { collectExternalFontFaceCss, collectFontFaceCss } from './foreign-object-font-embed';
import type {
	ExternalFontFaceCssResult,
	FontStyleDocumentLike,
	LinkStyleDocumentLike,
} from './foreign-object-font-embed';
/**
 * Assemble a self-contained SVG `foreignObject` document from a live slide
 * DOM subtree: the "does not depend on html2canvas's CSS subset" raster
 * path. Combines three already-shared primitives:
 *
 * - {@link inlineComputedStylesOnClone} (`foreign-object-style-inline.ts`):
 *   copies every computed style onto a detached clone, so custom properties,
 *   `backdrop-filter` and 3D transforms survive as concrete values the
 *   browser's own layout/paint engine (not html2canvas's) rasterises.
 * - {@link collectFontFaceCss} (`foreign-object-font-embed.ts`): reuses the
 *   deck's already-`data:`-URI `@font-face` CSS every binding injects for
 *   on-screen rendering.
 * - {@link embedImagesOnClone} (`foreign-object-image-embed.ts`): inlines any
 *   `blob:`/`http(s):` image reference that isn't already a `data:` URI.
 *
 * The result is split into a `bodyMarkup` (the `<defs>` + background rect +
 * `foreignObject` content, independent of output size) and a cheap
 * {@link wrapForeignObjectSvg} step that wraps it in an outer `<svg>` sized
 * for one export tile. Splitting the two means a tiled export pays the
 * (relatively) expensive clone/style/font/image work exactly once per slide,
 * then only re-wraps a viewBox window per tile.
 */
import { embedImagesOnClone, fetchAsDataUrl } from './foreign-object-image-embed';
import { inlineComputedStylesOnClone } from './foreign-object-style-inline';
import type { ComputedStyleReader } from './foreign-object-style-inline';
import { escapeSvgAttr } from './visual-effects';

/** Options for {@link buildForeignObjectSvgBody}. */
export interface ForeignObjectSvgBodyOptions {
	/** Natural (unscaled, CSS px) content width. */
	width: number;
	/** Natural (unscaled, CSS px) content height. */
	height: number;
	/** Optional solid background painted behind the foreignObject content. */
	backgroundColor?: string;
	/** Defaults to `window.getComputedStyle`; injectable for tests. */
	readComputedStyle?: ComputedStyleReader;
	/**
	 * The external font CSS ({@link collectExternalFontFaceCss}), when the caller already
	 * collected it for several bodies of one document; collected here otherwise.
	 */
	externalFonts?: ExternalFontFaceCssResult | Promise<ExternalFontFaceCssResult>;
}

/** The reusable, size-independent SVG body, plus whether every image embedded cleanly. */
export interface ForeignObjectSvgBody {
	/** `<defs>` (font-faces) + optional background rect + `<foreignObject>` markup. */
	bodyMarkup: string;
	/** `false` when at least one image reference could not be inlined as `data:`. */
	allEmbedded: boolean;
	naturalWidth: number;
	naturalHeight: number;
}

/**
 * Wrap CSS text in a CDATA section. A literal `]]>` inside the text (an
 * external stylesheet or a deck font name can carry one) would end the
 * section early and let the rest parse as markup, so it is split across two
 * sections, which XML joins back into the original text.
 */
export function cssCdata(css: string): string {
	return `<![CDATA[${css.replaceAll(']]>', ']]]]><![CDATA[>')}]]>`;
}

/**
 * Build the reusable SVG body for `element`: clone it, inline every computed
 * style, embed fonts and images, and return the `<defs>` + background +
 * `foreignObject` markup ready to be wrapped by {@link wrapForeignObjectSvg}.
 *
 * `element` must be attached to `doc` so `getComputedStyle` resolves real
 * values; the clone this function produces is fully detached and never
 * mutates the live tree.
 */
export async function buildForeignObjectSvgBody(
	element: HTMLElement,
	doc: Document & FontStyleDocumentLike & LinkStyleDocumentLike,
	options: ForeignObjectSvgBodyOptions,
): Promise<ForeignObjectSvgBody> {
	const prepared = await prepareForeignObjectClone(element, doc, options);
	return {
		bodyMarkup: foreignObjectBodyMarkup(
			new XMLSerializer().serializeToString(prepared.clone),
			prepared.fontFaceCss,
			options,
		),
		allEmbedded: prepared.allEmbedded,
		naturalWidth: options.width,
		naturalHeight: options.height,
	};
}

/** {@link prepareForeignObjectClone}'s result: the styled, self-contained clone, not yet serialised. */
export interface PreparedForeignObjectClone {
	/** The detached clone, every computed style inlined and every image embedded. */
	clone: HTMLElement;
	/** The `@font-face` CSS the body's `<defs>` carries. */
	fontFaceCss: string;
	/** `false` when at least one image or external font could not be inlined as `data:`. */
	allEmbedded: boolean;
}

/**
 * The expensive half of {@link buildForeignObjectSvgBody}: clone `element`, inline every computed
 * style, embed its images and collect the font CSS. A caller drawing several pictures of one
 * stage prepares it once and serialises variants of the clone with {@link foreignObjectBodyMarkup}.
 */
export async function prepareForeignObjectClone(
	element: HTMLElement,
	doc: Document & FontStyleDocumentLike & LinkStyleDocumentLike,
	options: Pick<ForeignObjectSvgBodyOptions, 'readComputedStyle' | 'externalFonts'> = {},
): Promise<PreparedForeignObjectClone> {
	const { readComputedStyle = (el) => window.getComputedStyle(el) } = options;

	// 3D views draw into a shadow-root canvas a clone cannot carry: let them
	// finish loading, then swap each cloned view for an <img> of its pixels
	// (while the trees are still aligned).
	await settleThreeViews(element);
	const clone = element.cloneNode(true) as HTMLElement;
	// Read computed styles at AUTHORED stroke widths, not the on-screen
	// one-device-pixel minimum the live stage applies.
	withAuthoredStrokeWidths(element, () =>
		inlineComputedStylesOnClone(element, clone, readComputedStyle),
	);
	snapshotThreeViewsIntoClone(element, clone);
	// Keep the trees aligned while copying styles, then omit editor-only nodes
	// before embedding resources. The live editor tree is never changed.
	prepareExportClone(clone);

	const [{ allEmbedded: imagesEmbedded }, externalFonts] = await Promise.all([
		embedImagesOnClone(clone),
		options.externalFonts ?? collectExternalFontFaceCss(doc, fetchAsDataUrl),
	]);
	const allEmbedded = imagesEmbedded && externalFonts.allEmbedded;
	const fontFaceCss = [collectFontFaceCss(doc), externalFonts.css].filter(Boolean).join('\n\n');
	return { clone, fontFaceCss, allEmbedded };
}

/**
 * The `<defs>` + background + `foreignObject` markup around `cloneMarkup` (a clone serialised
 * as XML), as {@link buildForeignObjectSvgBody} returns it.
 */
export function foreignObjectBodyMarkup(
	cloneMarkup: string,
	fontFaceCss: string,
	options: Pick<ForeignObjectSvgBodyOptions, 'width' | 'height' | 'backgroundColor'>,
): string {
	const { width, height, backgroundColor } = options;
	const defs = fontFaceCss
		? `<defs><style type="text/css">${cssCdata(fontFaceCss)}</style></defs>`
		: '';
	const bgRect = backgroundColor
		? `<rect width="${width}" height="${height}" fill="${escapeSvgAttr(backgroundColor)}" />`
		: '';
	return [
		defs,
		bgRect,
		`<foreignObject x="0" y="0" width="${width}" height="${height}">`,
		`<div xmlns="http://www.w3.org/1999/xhtml" style="width:${width}px;height:${height}px;overflow:hidden;">`,
		// This becomes XML, not HTML: retain nested SVG namespaces and close void elements.
		cloneMarkup,
		'</div>',
		'</foreignObject>',
	].join('');
}

/** A viewBox window into the body's natural coordinate space, for one output tile. */
export interface ForeignObjectTileWindow {
	/** Left edge of the window, in the body's natural (unscaled) coordinates. */
	viewBoxX: number;
	/** Top edge of the window, in the body's natural (unscaled) coordinates. */
	viewBoxY: number;
	/** Window width, in the body's natural (unscaled) coordinates. */
	viewBoxWidth: number;
	/** Window height, in the body's natural (unscaled) coordinates. */
	viewBoxHeight: number;
	/** Output raster width in device pixels for this tile. */
	outputWidth: number;
	/** Output raster height in device pixels for this tile. */
	outputHeight: number;
}

/**
 * Wrap a {@link ForeignObjectSvgBody.bodyMarkup} in an outer `<svg>` sized and
 * windowed for one output tile. For the common non-tiled case, pass a window
 * covering the full natural size with `outputWidth`/`outputHeight` at the
 * desired export scale.
 *
 * SECURITY: `bodyMarkup` carries DOM-derived text (cloned slide content, font
 * CSS), so the returned string must only reach a non-executing sink: an
 * `<img>` `data:image/svg+xml` source ({@link foreignObjectSvgToDataUrl} in
 * `export/rasterize-foreign-object.ts`, the only production caller) or an
 * XML-mode `DOMParser().parseFromString(xml, 'image/svg+xml')`. Both render
 * SVG in "image" context, which never executes embedded scripts. Never pass
 * this string to `innerHTML`/`outerHTML`, `document.write`, or an HTML-mode
 * parse.
 *
 * Every deck-derived value in `bodyMarkup` is already escaped by
 * {@link buildForeignObjectSvgBody}: slide text and attributes go through
 * `XMLSerializer`, the background colour through `escapeSvgAttr`, and font CSS
 * through {@link cssCdata}. The tile numbers are coerced with `Number`.
 */
export function wrapForeignObjectSvg(bodyMarkup: string, tile: ForeignObjectTileWindow): string {
	const { viewBoxX, viewBoxY, viewBoxWidth, viewBoxHeight, outputWidth, outputHeight } = tile;
	return (
		`<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" ` +
		`width="${Number(outputWidth)}" height="${Number(outputHeight)}" ` +
		`viewBox="${Number(viewBoxX)} ${Number(viewBoxY)} ${Number(viewBoxWidth)} ${Number(viewBoxHeight)}">${bodyMarkup}</svg>`
	);
}
