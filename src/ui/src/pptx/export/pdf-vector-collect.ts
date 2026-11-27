/**
 * Read a mounted slide stage (any binding's) into a display list for vector PDF export. An
 * allow-list decides what is vector: element boxes (`pdf-vector-box.ts`), pictures, inline SVG
 * the shared converter accepts, and text. Anything else (filters, blend modes, masks, 3-D,
 * canvases, shadows, unusual fills) is captured as a picture of just that element, in paint
 * order, with its text kept as invisible, searchable text. The walk runs with the stage's CSS
 * transforms neutralised so text boxes are measured unrotated; each element's matrix is
 * reapplied when painting.
 */
import type { PdfFontSet } from 'ooxml-core/pdf/write';
import { decodePdfImage, drawableToPdfImage } from '../../export/pdf-image';
import type { PdfImageResource } from 'ooxml-core/pdf/write';
import { withAuthoredStrokeWidths } from '../render/device-pixel-stroke';
import { isExportIgnoredElement } from '../render/export-clone';
import {
	PdfDiagnosticSet,
	type PdfExportDiagnostic,
	type PdfExportDiagnosticCode,
} from './pdf-export-diagnostics';
import { paintBox } from './pdf-vector-box';
import { videoPicture } from './pdf-vector-media';
import { FONTS_TIMEOUT_MS, IMAGE_TIMEOUT_MS, settleWithin } from './pdf-vector-timeout';
import {
	boxClip,
	contentBox,
	effectPad,
	fitImage,
	hasOwnText,
	hasText,
	inlineSvgItem,
	isCollapsedCell,
	MEDIA_REASON,
	MEDIA_TAGS,
	OPACITY_NOTE,
	paddingBox,
	paintedBounds,
	placeholderImage,
	SEARCHABLE_NOTE,
	subtreeRasterReason,
	SVG_REASON,
	TEXT_SHADOW_REASON,
	THREE_D_REASON,
	type Box,
} from './pdf-vector-collect-parts';
import { clipPathData } from './pdf-vector-css';
import type { DisplayItem, ImageItem, Matrix, PdfClip } from './pdf-vector-display-list';
import { IDENTITY, multiply } from './pdf-vector-display-list';
import type { IsolatedRasterMode } from './pdf-vector-raster';
import { rasterizeIsolatedBatch, type IsolatedRasterBatchOptions } from './pdf-vector-raster-batch';
import { MAX_TEXT_TOKENS, measureTextNode } from './pdf-vector-text';
import {
	neutraliseTransforms,
	type ElementTransforms,
	type StageMeasure,
} from './pdf-vector-transforms';
import type { RasterStrategyOptions } from './rasterize-element-strategy';

export interface CollectSlideOptions {
	canvasWidth: number;
	canvasHeight: number;
	fonts?: PdfFontSet;
	/** Device pixels per CSS px for element pictures. Default 2. */
	rasterScale?: number;
	html2canvasFallback?: RasterStrategyOptions['html2canvasFallback'];
	/** Editor chrome a binding marks its own way (beyond `data-export-ignore`). */
	ignoreElement?: (element: Element) => boolean;
	/** The most element pictures per slide before the whole slide is one picture. Default 32. */
	maxRasters?: number;
	/** The document's external font CSS, shared by every slide of one export. */
	externalFonts?: IsolatedRasterBatchOptions['externalFonts'];
}

export interface CollectedSlide {
	items: DisplayItem[];
	diagnostics: PdfExportDiagnostic[];
	/** Elements drawn as pictures of just themselves (effects, media frames). */
	pictures: number;
	/** Set when the whole slide must be drawn as one picture (with `items` as invisible text). */
	picture?: PdfExportDiagnosticCode;
}

interface Pending {
	item: ImageItem;
	load: () => Promise<PdfImageResource | undefined>;
	fit?: 'fill' | 'cover' | 'contain';
	failure: PdfExportDiagnosticCode;
	/** A media frame (canvas, video): counts as an element drawn as a picture. */
	frame?: boolean;
}

interface Context {
	matrix: Matrix;
	clips: readonly PdfClip[];
	alpha: number;
	invisible: boolean;
}

/** Collect `stage` (mounted and laid out) for a `canvasWidth` x `canvasHeight` slide. */
export async function collectSlide(
	stage: HTMLElement,
	options: CollectSlideOptions,
): Promise<CollectedSlide> {
	const doc = stage.ownerDocument;
	const view = doc.defaultView!;
	const items: DisplayItem[] = [];
	const notes = new PdfDiagnosticSet();
	const pending: Pending[] = [];
	const rasters: { item: ImageItem; target: Element; mode: IsolatedRasterMode }[] = [];
	const svgImages: { hrefs: string[]; map: Map<string, PdfImageResource> }[] = [];
	let tokens = 0;
	let picture: PdfExportDiagnosticCode | undefined;
	let pictures = 0;
	const W = options.canvasWidth,
		H = options.canvasHeight;
	// The stage's own layout size: a zoomed canvas lays the slide out at another size.
	const layoutW = stage.offsetWidth || W,
		layoutH = stage.offsetHeight || H;
	const sx = W / layoutW,
		sy = H / layoutH;
	let measure: StageMeasure = { originX: 0, originY: 0, scale: 1 };
	let transforms: ElementTransforms = new Map();

	const boxOf = (el: Element): Box => {
		const rect = el.getBoundingClientRect();
		const k = measure.scale;
		return {
			x: (rect.left - measure.originX) / k,
			y: (rect.top - measure.originY) / k,
			w: rect.width / k,
			h: rect.height / k,
		};
	};

	/** A picture of `target`, over its painted bounds plus `pad`, at this point of the order. */
	const raster = (
		target: Element,
		mode: IsolatedRasterMode,
		matrix: Matrix,
		box: Box,
		pad: number,
		reason: PdfExportDiagnosticCode,
	) => {
		notes.add(reason);
		const bounds = paintedBounds(matrix, box, pad, W, H);
		if (!bounds) return;
		const item = placeholderImage(IDENTITY, [], 1, bounds);
		items.push(item);
		rasters.push({ item, target, mode });
	};

	const walk = (el: Element, ctx: Context) => {
		if (picture || isExportIgnoredElement(el) || options.ignoreElement?.(el)) return;
		const style = view.getComputedStyle(el);
		if (style.display === 'none') return;
		const box = boxOf(el);
		let matrix = ctx.matrix;
		const transform = transforms.get(el);
		let reason: PdfExportDiagnosticCode | undefined;
		if (transform?.local) matrix = multiply(matrix, transform.local);
		else if (transform && !ctx.invisible) reason = THREE_D_REASON;
		const opacity = Number.parseFloat(style.opacity);
		const alpha = ctx.alpha * (Number.isFinite(opacity) ? opacity : 1);
		if (alpha <= 0.001) return;
		if (opacity < 1 && el.childElementCount > 1 && !ctx.invisible) notes.add(OPACITY_NOTE);
		const clip = clipPathData(style.clipPath, box.w, box.h);
		if (!ctx.invisible) reason ??= subtreeRasterReason(el, style, clip !== undefined);
		if (reason) {
			raster(el, 'subtree', matrix, box, effectPad(reason, box), reason);
			if (hasText(el)) notes.add(SEARCHABLE_NOTE);
			walkChildren(el, { matrix, clips: ctx.clips, alpha, invisible: true }, style);
			return;
		}
		const visible = style.visibility === 'visible';
		const clipMatrix = multiply(matrix, [1, 0, 0, 1, box.x, box.y]);
		const clips = clip
			? [...ctx.clips, { matrix: clipMatrix, d: clip.d, evenOdd: clip.evenOdd }]
			: ctx.clips;
		if (visible && !ctx.invisible) {
			// The stage's own shadow is editor chrome (the page on the canvas), not slide content.
			const shadow =
				el === stage
					? 'none'
					: (el.getAttribute('data-export-original-box-shadow') ?? style.boxShadow);
			const geometry = { x: box.x, y: box.y, width: box.w, height: box.h };
			const collapsed = isCollapsedCell(el, style, view);
			const paint = paintBox(style, geometry, matrix, clips, alpha, shadow, collapsed);
			if (paint.raster) {
				const own = hasOwnText(el);
				raster(
					el,
					own ? 'subtree' : 'self',
					matrix,
					box,
					effectPad(paint.raster, box),
					paint.raster,
				);
				if (own) {
					notes.add(SEARCHABLE_NOTE);
					walkChildren(el, { matrix, clips, alpha, invisible: true }, style);
					return;
				}
			} else {
				items.push(...paint.items);
				if (paint.image) {
					const { href, fit, item } = paint.image;
					const failure = 'backgroundPictureMissing';
					pending.push({ item, fit, load: () => decodePdfImage(doc, href), failure });
				}
			}
		}
		const tag = el.localName;
		const content = contentBox(style, box);
		if (tag === 'img' || MEDIA_TAGS.has(tag)) {
			if (!visible || ctx.invisible || !(content.w > 0 && content.h > 0)) return;
			const media = el as HTMLImageElement & HTMLCanvasElement & HTMLVideoElement;
			const natural =
				tag === 'img' ? [media.naturalWidth, media.naturalHeight] : [media.width, media.height];
			if (tag !== 'img') notes.add(MEDIA_REASON);
			const item = placeholderImage(matrix, [...clips, boxClip(matrix, content)], alpha, content);
			items.push(item);
			const fit =
				style.objectFit === 'cover' ? 'cover' : style.objectFit === 'contain' ? 'contain' : 'fill';
			const load = () =>
				tag === 'video'
					? videoPicture(doc, media)
					: drawableToPdfImage(doc, media, natural[0] ?? 0, natural[1] ?? 0);
			pending.push({
				item,
				fit,
				load,
				failure: 'pictureMissing',
				...(tag !== 'img' ? { frame: true } : {}),
			});
			return;
		}
		if (tag === 'svg') {
			if (ctx.invisible) {
				walkChildren(el, { ...ctx, matrix, clips, alpha }, style);
				return;
			}
			const svg = inlineSvgItem(el as SVGSVGElement, style, matrix, clips, alpha, content);
			if (!svg) {
				raster(el, 'subtree', matrix, box, 24, SVG_REASON);
				return;
			}
			svgImages.push({ hrefs: svg.hrefs, map: svg.images });
			items.push(svg.item);
			return;
		}
		const overflow = style.overflowX !== 'visible' || style.overflowY !== 'visible';
		const childClips = overflow ? [...clips, boxClip(matrix, paddingBox(style, box))] : clips;
		walkChildren(el, { matrix, clips: childClips, alpha, invisible: ctx.invisible }, style);
	};

	const walkChildren = (el: Element, ctx: Context, style: CSSStyleDeclaration) => {
		if (el.localName === 'svg' && ctx.invisible) return;
		const children = Array.from(el.childNodes).map((node, index) => {
			let z = 0;
			if (node.nodeType === 1) {
				const cs = view.getComputedStyle(node as Element);
				if (cs.position !== 'static' && cs.zIndex !== 'auto')
					z = Number.parseInt(cs.zIndex, 10) || 0;
			}
			return { node, z, index };
		});
		children.sort((a, b) => a.z - b.z || a.index - b.index);
		let shadowDone = false;
		// One picture of `el` covers all of its text nodes that need one.
		let textPicture = false;
		for (const { node } of children) {
			if (picture) return;
			if (node.nodeType === 1) {
				walk(node as Element, ctx);
				continue;
			}
			if (node.nodeType !== 3 || style.visibility !== 'visible') continue;
			const text = node as Text;
			const shadowed = style.textShadow && style.textShadow !== 'none';
			if (!shadowDone && !ctx.invisible && shadowed && text.data.trim()) {
				shadowDone = true;
				raster(el, 'shadow', ctx.matrix, boxOf(el), 24, TEXT_SHADOW_REASON);
			}
			const result = measureTextNode(
				text,
				style,
				{
					matrix: ctx.matrix,
					clips: ctx.clips,
					alpha: ctx.alpha,
					...measure,
					...(options.fonts ? { fonts: options.fonts } : {}),
					invisible: ctx.invisible,
				},
				MAX_TEXT_TOKENS - tokens,
			);
			tokens += result.tokens;
			if (tokens > MAX_TEXT_TOKENS) {
				picture = 'slideTooMuchText';
				return;
			}
			if (result.raster) {
				if (!textPicture) raster(el, 'subtree', ctx.matrix, boxOf(el), 4, result.raster);
				else notes.add(result.raster);
				textPicture = true;
				items.push(...result.items.map((item) => ({ ...item, invisible: true })));
			} else items.push(...result.items);
		}
	};

	if (doc.fonts) await settleWithin(doc.fonts.ready, FONTS_TIMEOUT_MS, undefined);
	withAuthoredStrokeWidths(stage, () =>
		neutraliseTransforms(stage, view, (found, stageMeasure) => {
			transforms = found;
			measure = stageMeasure;
			walk(stage, { matrix: [sx, 0, 0, sy, 0, 0], clips: [], alpha: 1, invisible: false });
		}),
	);
	if (!picture && rasters.length > (options.maxRasters ?? 32)) picture = 'slideTooManyEffects';
	if (picture) {
		return {
			items: items.flatMap((item) => (item.kind === 'text' ? [{ ...item, invisible: true }] : [])),
			diagnostics: [{ code: picture }, { code: SEARCHABLE_NOTE }],
			pictures: 0,
			picture,
		};
	}
	const failed = new Set<DisplayItem>();
	for (const request of pending) {
		const image = await settleWithin(request.load(), IMAGE_TIMEOUT_MS, undefined);
		if (!image) {
			failed.add(request.item);
			notes.add(request.failure);
			continue;
		}
		if (request.frame) pictures++;
		request.item.image = image;
		if (request.fit && request.fit !== 'fill') fitImage(request.item, image, request.fit);
	}
	for (const { hrefs, map } of svgImages)
		for (const href of hrefs.slice(0, 500)) {
			const image = await settleWithin(decodePdfImage(doc, href), IMAGE_TIMEOUT_MS, undefined);
			if (image) map.set(href, image);
		}
	const images = await rasterizeIsolatedBatch(
		stage,
		rasters.map(({ item, target, mode }) => ({
			target,
			mode,
			rect: { x: item.x / sx, y: item.y / sy, width: item.width / sx, height: item.height / sy },
		})),
		layoutW,
		layoutH,
		{
			scale: (options.rasterScale ?? 2) * sx,
			...(options.html2canvasFallback ? { html2canvasFallback: options.html2canvasFallback } : {}),
			...(options.externalFonts ? { externalFonts: options.externalFonts } : {}),
		},
	);
	rasters.forEach(({ item }, i) => {
		const image = images[i];
		if (image) {
			item.image = image;
			pictures++;
		} else {
			failed.add(item);
			notes.add('effectMissing');
		}
	});
	return {
		items: items.filter((item) => !failed.has(item)),
		diagnostics: notes.toArray(),
		pictures,
	};
}
