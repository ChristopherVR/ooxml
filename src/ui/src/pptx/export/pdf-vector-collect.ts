/**
 * Read a mounted slide stage (any binding's) into a display list for vector PDF export. An
 * allow-list decides what is vector: element boxes (`pdf-vector-box.ts`), pictures, inline SVG
 * the shared converter accepts, and text. Anything else (filters, blend modes, masks, 3-D,
 * canvases, shadows, unusual fills) is captured as a picture of just that element, in paint
 * order, with its text kept as invisible, searchable text. The walk runs with the stage's CSS
 * transforms neutralised so text boxes are measured unrotated; each element's matrix is
 * reapplied when painting.
 */
import type { PdfFontSet } from '../../export/pdf-embedded-font';
import { decodePdfImage, drawableToPdfImage } from '../../export/pdf-image';
import { collectSvgImages, drawSvgSubtree } from '../../export/pdf-svg';
import { PdfFallback } from '../../export/pdf-svg-style';
import { PdfVectorPage, type PdfImageResource } from '../../export/pdf-vector';
import { isExportIgnoredElement } from '../render/export-clone';
import { withAuthoredStrokeWidths } from '../render/device-pixel-stroke';
import { paintBox, SHADOW_REASON } from './pdf-vector-box';
import { clipPathData, roundedRectPath } from './pdf-vector-css';
import type { DisplayItem, ImageItem, Matrix, PdfClip } from './pdf-vector-display-list';
import { IDENTITY, multiply } from './pdf-vector-display-list';
import { rasterizeIsolated, type IsolatedRasterMode } from './pdf-vector-raster';
import type { RasterStrategyOptions } from './rasterize-element-strategy';
import { MAX_TEXT_TOKENS, measureTextNode } from './pdf-vector-text';
import {
	neutraliseTransforms,
	type ElementTransforms,
	type StageMeasure,
} from './pdf-vector-transforms';
import { normalizeInlineSvg, viewBoxMatrix } from './pdf-vector-svg';

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
}

export interface CollectedSlide {
	items: DisplayItem[];
	diagnostics: string[];
	/** Set when the whole slide must be drawn as one picture (with `items` as invisible text). */
	picture?: string;
}

const FILTER_REASON = 'Soft edges, blur, glow and other filter effects are drawn as pictures.';
const BLEND_REASON = 'Blend modes and masks are drawn as pictures.';
const THREE_D_REASON = '3-D rotation and perspective are drawn as pictures.';
const MEDIA_REASON =
	'Video, 3-D models and canvas drawings are drawn as pictures of their current frame.';
const TEXT_SHADOW_REASON = 'Text shadows are drawn as pictures under the searchable text.';
const SEARCHABLE_NOTE = 'Text inside pictures stays searchable as invisible text.';
const OPACITY_NOTE = 'Transparency set on a group is applied to each of its objects separately.';
const SVG_REASON = 'Drawings this export cannot write as vector graphics are drawn as pictures.';

interface Pending {
	item: ImageItem;
	load: () => Promise<PdfImageResource | undefined>;
	fit?: 'fill' | 'cover' | 'contain';
	failure: string;
}

interface Context {
	matrix: Matrix;
	clips: readonly PdfClip[];
	alpha: number;
	invisible: boolean;
}

const px = (value: string) => Number.parseFloat(value) || 0;
const MEDIA_TAGS = new Set(['canvas', 'video']);
const OPAQUE_TAGS = new Set(['iframe', 'object', 'embed', 'math']);

/** Collect `stage` (mounted, laid out, `canvasWidth` x `canvasHeight` CSS px). */
export async function collectSlide(
	stage: HTMLElement,
	options: CollectSlideOptions,
): Promise<CollectedSlide> {
	const doc = stage.ownerDocument;
	const view = doc.defaultView!;
	const items: DisplayItem[] = [];
	const notes = new Set<string>();
	const pending: Pending[] = [];
	const rasters: { item: ImageItem; target: Element; mode: IsolatedRasterMode }[] = [];
	const svgImages: { hrefs: string[]; map: Map<string, PdfImageResource> }[] = [];
	let tokens = 0;
	let picture: string | undefined;
	const W = options.canvasWidth,
		H = options.canvasHeight;
	// The stage's own layout size: a zoomed canvas lays the slide out at another size.
	const layoutW = stage.offsetWidth || W,
		layoutH = stage.offsetHeight || H;
	const sx = W / layoutW,
		sy = H / layoutH;

	const placeholder = (
		matrix: Matrix,
		clips: readonly PdfClip[],
		alpha: number,
		x: number,
		y: number,
		w: number,
		h: number,
	): ImageItem => ({
		kind: 'image',
		matrix,
		clips,
		alpha,
		x,
		y,
		width: w,
		height: h,
		image: { width: 0, height: 0, data: new Uint8Array() },
	});

	/** A picture of `target`, over its painted bounds plus `pad`, at this point of the order. */
	const raster = (
		target: Element,
		mode: IsolatedRasterMode,
		matrix: Matrix,
		box: { x: number; y: number; w: number; h: number },
		pad: number,
		reason: string,
	) => {
		notes.add(reason);
		const corners = [
			[box.x, box.y],
			[box.x + box.w, box.y],
			[box.x, box.y + box.h],
			[box.x + box.w, box.y + box.h],
		].map(([cx, cy]) => [
			matrix[0] * cx! + matrix[2] * cy! + matrix[4],
			matrix[1] * cx! + matrix[3] * cy! + matrix[5],
		]);
		const x0 = Math.max(0, Math.floor(Math.min(...corners.map((c) => c[0]!)) - pad));
		const y0 = Math.max(0, Math.floor(Math.min(...corners.map((c) => c[1]!)) - pad));
		const x1 = Math.min(W, Math.ceil(Math.max(...corners.map((c) => c[0]!)) + pad));
		const y1 = Math.min(H, Math.ceil(Math.max(...corners.map((c) => c[1]!)) + pad));
		if (!(x1 > x0 && y1 > y0)) return;
		const item = placeholder(IDENTITY, [], 1, x0, y0, x1 - x0, y1 - y0);
		items.push(item);
		rasters.push({ item, target, mode });
	};

	let measure: StageMeasure = { originX: 0, originY: 0, scale: 1 };
	let transforms: ElementTransforms = new Map();
	const boxOf = (el: Element) => {
		const rect = el.getBoundingClientRect();
		const k = measure.scale;
		return {
			x: (rect.left - measure.originX) / k,
			y: (rect.top - measure.originY) / k,
			w: rect.width / k,
			h: rect.height / k,
		};
	};

	const walk = (el: Element, ctx: Context) => {
		if (picture || isExportIgnoredElement(el) || options.ignoreElement?.(el)) return;
		const style = view.getComputedStyle(el);
		if (style.display === 'none') return;
		const box = boxOf(el);
		let matrix = ctx.matrix;
		const transform = transforms.get(el);
		let reason: string | undefined;
		if (transform) {
			if (transform.local) matrix = multiply(matrix, transform.local);
			else reason = THREE_D_REASON;
		}
		const opacity = Number.parseFloat(style.opacity);
		const alpha = ctx.alpha * (Number.isFinite(opacity) ? opacity : 1);
		if (alpha <= 0.001) return;
		const tag = el.localName;
		if (opacity < 1 && el.childElementCount > 1 && !ctx.invisible) notes.add(OPACITY_NOTE);
		const clip = clipPathData(style.clipPath, box.w, box.h);
		if (!ctx.invisible && !reason) {
			if (style.filter && style.filter !== 'none') reason = FILTER_REASON;
			else if ((style.mixBlendMode && style.mixBlendMode !== 'normal') || hasMask(style))
				reason = BLEND_REASON;
			else if (clip === undefined) reason = BLEND_REASON;
			else if (OPAQUE_TAGS.has(tag) || (el as HTMLElement).shadowRoot) reason = MEDIA_REASON;
			else if (
				style.backgroundClip === 'text' ||
				style.getPropertyValue('-webkit-background-clip') === 'text'
			)
				reason = FILL_TEXT_REASON;
		}
		if (reason) {
			raster(
				el,
				'subtree',
				matrix,
				box,
				reason === FILTER_REASON || reason === THREE_D_REASON
					? Math.min(200, 24 + 0.25 * Math.max(box.w, box.h))
					: 2,
				reason,
			);
			if (hasText(el)) notes.add(SEARCHABLE_NOTE);
			walkChildren(el, { matrix, clips: ctx.clips, alpha, invisible: true }, style);
			return;
		}
		const visible = style.visibility === 'visible';
		let clips = ctx.clips;
		if (clip)
			clips = [
				...clips,
				{ matrix: multiply(matrix, [1, 0, 0, 1, box.x, box.y]), d: clip.d, evenOdd: clip.evenOdd },
			];
		if (visible && !ctx.invisible) {
			// The stage's own shadow is editor chrome (the page on the canvas), not slide content.
			const shadow =
				el === stage
					? 'none'
					: (el.getAttribute('data-export-original-box-shadow') ?? style.boxShadow);
			const paint = paintBox(
				style,
				{ x: box.x, y: box.y, width: box.w, height: box.h },
				matrix,
				clips,
				alpha,
				shadow,
				style.display === 'table-cell' &&
					!!el.closest('table') &&
					view.getComputedStyle(el.closest('table')!).borderCollapse === 'collapse',
			);
			if (paint.raster) {
				const own = hasOwnText(el);
				raster(
					el,
					own ? 'subtree' : 'self',
					matrix,
					box,
					paint.raster === SHADOW_REASON ? Math.min(200, 24 + 0.25 * Math.max(box.w, box.h)) : 2,
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
					pending.push({
						item,
						fit,
						load: () => decodePdfImage(doc, href),
						failure: 'A background picture could not be read and is left out.',
					});
				}
			}
		}
		const content = {
			x: box.x + px(style.borderLeftWidth) + px(style.paddingLeft),
			y: box.y + px(style.borderTopWidth) + px(style.paddingTop),
			w:
				box.w -
				px(style.borderLeftWidth) -
				px(style.borderRightWidth) -
				px(style.paddingLeft) -
				px(style.paddingRight),
			h:
				box.h -
				px(style.borderTopWidth) -
				px(style.borderBottomWidth) -
				px(style.paddingTop) -
				px(style.paddingBottom),
		};
		if (tag === 'img' || MEDIA_TAGS.has(tag)) {
			if (!visible || ctx.invisible || !(content.w > 0 && content.h > 0)) return;
			const media = el as HTMLImageElement | HTMLCanvasElement | HTMLVideoElement;
			const natural =
				tag === 'img'
					? [(media as HTMLImageElement).naturalWidth, (media as HTMLImageElement).naturalHeight]
					: tag === 'video'
						? [(media as HTMLVideoElement).videoWidth, (media as HTMLVideoElement).videoHeight]
						: [(media as HTMLCanvasElement).width, (media as HTMLCanvasElement).height];
			if (tag !== 'img') notes.add(MEDIA_REASON);
			const item = placeholder(
				matrix,
				[...clips, { matrix, d: roundedRectPath(content.x, content.y, content.w, content.h) }],
				alpha,
				content.x,
				content.y,
				content.w,
				content.h,
			);
			items.push(item);
			const fit =
				style.objectFit === 'cover' ? 'cover' : style.objectFit === 'contain' ? 'contain' : 'fill';
			pending.push({
				item,
				fit,
				load: () => drawableToPdfImage(doc, media, natural[0] ?? 0, natural[1] ?? 0),
				failure: 'A picture could not be read and is left out.',
			});
			return;
		}
		if (tag === 'svg') {
			if (ctx.invisible) {
				walkChildren(el, { ...ctx, matrix, clips, alpha }, style);
				return;
			}
			const copy = normalizeInlineSvg(el as SVGSVGElement);
			const svgMatrix = multiply(
				matrix,
				multiply([1, 0, 0, 1, content.x, content.y], viewBoxMatrix(el, content.w, content.h)),
			);
			const svgClips =
				style.overflow === 'visible'
					? clips
					: [...clips, { matrix, d: roundedRectPath(content.x, content.y, content.w, content.h) }];
			let ok = !copy.querySelector('[filter], [mask], foreignObject, pattern');
			if (ok) {
				try {
					drawSvgSubtree(
						new PdfVectorPage(10, 10),
						copy,
						new Map(
							collectSvgImages(copy).map((href) => [
								href,
								{ width: 1, height: 1, data: new Uint8Array(3) },
							]),
						),
					);
				} catch (error) {
					if (!(error instanceof PdfFallback)) throw error;
					ok = false;
				}
			}
			if (!ok) {
				raster(el, 'subtree', matrix, box, 24, SVG_REASON);
				return;
			}
			const map = new Map<string, PdfImageResource>();
			svgImages.push({ hrefs: collectSvgImages(copy), map });
			items.push({
				kind: 'svg',
				matrix: svgMatrix,
				clips: svgClips,
				alpha,
				svg: copy,
				images: map,
			});
			return;
		}
		const childClips =
			style.overflowX !== 'visible' || style.overflowY !== 'visible'
				? [
						...clips,
						{
							matrix,
							d: roundedRectPath(
								box.x + px(style.borderLeftWidth),
								box.y + px(style.borderTopWidth),
								box.w - px(style.borderLeftWidth) - px(style.borderRightWidth),
								box.h - px(style.borderTopWidth) - px(style.borderBottomWidth),
							),
						},
					]
				: clips;
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
		for (const { node } of children) {
			if (picture) return;
			if (node.nodeType === 1) walk(node as Element, ctx);
			else if (node.nodeType === 3 && style.visibility === 'visible') {
				if (
					!shadowDone &&
					!ctx.invisible &&
					style.textShadow &&
					style.textShadow !== 'none' &&
					(node as Text).data.trim()
				) {
					shadowDone = true;
					raster(el, 'shadow', ctx.matrix, boxOf(el), 24, TEXT_SHADOW_REASON);
				}
				const result = measureTextNode(
					node as Text,
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
					picture =
						'The slide has more text than this export reads, so it is drawn as one picture.';
					return;
				}
				if (result.raster) {
					raster(el, 'subtree', ctx.matrix, boxOf(el), 4, result.raster);
					items.push(...result.items.map((item) => ({ ...item, invisible: true })));
				} else items.push(...result.items);
			}
		}
	};

	await doc.fonts?.ready;
	withAuthoredStrokeWidths(stage, () =>
		neutraliseTransforms(stage, view, (found, stageMeasure) => {
			transforms = found;
			measure = stageMeasure;
			walk(stage, { matrix: [sx, 0, 0, sy, 0, 0], clips: [], alpha: 1, invisible: false });
		}),
	);
	const limit = options.maxRasters ?? 32;
	if (!picture && rasters.length > limit)
		picture = 'The slide has too many effects to draw separately, so it is drawn as one picture.';
	if (picture) {
		return {
			items: items.flatMap((item) => (item.kind === 'text' ? [{ ...item, invisible: true }] : [])),
			diagnostics: [picture, SEARCHABLE_NOTE],
			picture,
		};
	}
	const failed = new Set<DisplayItem>();
	for (const request of pending) {
		const image = await request.load();
		if (!image) {
			failed.add(request.item);
			notes.add(request.failure);
			continue;
		}
		request.item.image = image;
		if (request.fit && request.fit !== 'fill') fitImage(request.item, image, request.fit);
	}
	for (const { hrefs, map } of svgImages)
		for (const href of hrefs.slice(0, 500)) {
			const image = await decodePdfImage(doc, href);
			if (image) map.set(href, image);
		}
	for (const { item, target, mode } of rasters) {
		const image = await rasterizeIsolated(
			stage,
			target,
			mode,
			{ x: item.x / sx, y: item.y / sy, width: item.width / sx, height: item.height / sy },
			layoutW,
			layoutH,
			{
				scale: (options.rasterScale ?? 2) * sx,
				...(options.html2canvasFallback
					? { html2canvasFallback: options.html2canvasFallback }
					: {}),
			},
		);
		if (image) item.image = image;
		else {
			failed.add(item);
			notes.add('An effect could not be captured as a picture and is left out.');
		}
	}
	return { items: items.filter((item) => !failed.has(item)), diagnostics: [...notes] };
}

const FILL_TEXT_REASON = 'Text filled with a gradient or picture is drawn as a picture.';

function hasMask(style: CSSStyleDeclaration): boolean {
	const mask = style.getPropertyValue('mask-image') || style.getPropertyValue('-webkit-mask-image');
	const backdrop =
		style.getPropertyValue('backdrop-filter') || style.getPropertyValue('-webkit-backdrop-filter');
	return (!!mask && mask !== 'none') || (!!backdrop && backdrop !== 'none');
}

const hasText = (el: Element) => !!el.textContent?.trim();
const hasOwnText = (el: Element) =>
	Array.from(el.childNodes).some((node) => node.nodeType === 3 && !!(node as Text).data.trim());

/** Fit a decoded picture into its box like `object-fit` / `background-size`. */
function fitImage(item: ImageItem, image: PdfImageResource, fit: 'cover' | 'contain'): void {
	const ratio = image.width / image.height;
	if (!(ratio > 0)) return;
	const boxRatio = item.width / item.height;
	const wide = fit === 'cover' ? ratio > boxRatio : ratio < boxRatio;
	const w = wide ? item.height * ratio : item.width;
	const h = wide ? item.height : item.width / ratio;
	item.x += (item.width - w) / 2;
	item.y += (item.height - h) / 2;
	item.width = w;
	item.height = h;
}
