/**
 * Draw several elements of one mounted slide as pictures of just themselves, sharing one styled
 * clone of the stage. Cloning the stage, inlining every computed style and embedding its images
 * is the expensive part of a capture and is the same for every element of a slide, so it is done
 * once and serialised as a template (`pdf-vector-raster-template.ts`); each picture then only
 * writes which elements are visible (`isolatedMarkup`) and rasterises the element's own region.
 * The pictures are the ones {@link rasterizeIsolated} draws one at a time, which stays the path
 * for a capture this cannot take (a resource that would not embed, a target the clone lost, a
 * failed draw), so the html2canvas fallback still sees the live stage isolated as it always has.
 */
import { drawableToPdfImage } from '../../export/pdf-image';
import type { PdfImageResource } from 'ooxml-core/pdf/write';
import { collectExternalFontFaceCss } from '../render/foreign-object-font-embed';
import type {
	ExternalFontFaceCssResult,
	FontStyleDocumentLike,
	LinkStyleDocumentLike,
} from '../render/foreign-object-font-embed';
import { fetchAsDataUrl } from '../render/foreign-object-image-embed';
import type { ComputedStyleReader } from '../render/foreign-object-style-inline';
import {
	foreignObjectBodyMarkup,
	prepareForeignObjectClone,
} from '../render/foreign-object-svg-document';
import {
	rasterizeIsolated,
	type IsolatedRasterMode,
	type IsolatedRasterOptions,
} from './pdf-vector-raster';
import { isolatedMarkup, type IsolatedTemplate } from './pdf-vector-raster-isolate';
import { markDeclaredVisibility, shadowTargetStyle } from './pdf-vector-raster-live';
import { buildIsolatedTemplate } from './pdf-vector-raster-template';
import { RASTER_TIMEOUT_MS, settleWithin } from './pdf-vector-timeout';
import type { RasterSourceRect } from './rasterize-element-strategy';
import { rasterizeWindow } from './rasterize-element-strategy';

/** One picture to draw: `target` in `mode`, over `rect` (stage px, as painted). */
export interface IsolatedRasterRequest {
	target: Element;
	mode: IsolatedRasterMode;
	rect: RasterSourceRect;
}

export interface IsolatedRasterBatchOptions extends IsolatedRasterOptions {
	/** Injectable for tests. */
	readComputedStyle?: ComputedStyleReader;
}

/** Marks the raster targets on the live stage while it is cloned; removed straight after. */
export const RASTER_TARGET_ATTRIBUTE = 'data-pdf-raster-target';

/** One stage's styled clone, serialised once as a template every picture of it is cut from. */
export interface IsolatedRasterBase {
	template: IsolatedTemplate;
	/** Clone element index of each target, by the target's tag. */
	targets: Map<string, number>;
	fontFaceCss: string;
	allEmbedded: boolean;
}

/**
 * Clone `stage` once, with each `targets` element tagged so its clone can be found, within
 * {@link RASTER_TIMEOUT_MS}. The live stage is tagged only while it is cloned: the tags come off
 * before this returns, even when the clone stalls. `undefined` when the clone did not finish.
 */
export async function prepareIsolatedRasterBase(
	stage: HTMLElement,
	targets: ReadonlyMap<Element, string>,
	options: Pick<IsolatedRasterBatchOptions, 'externalFonts' | 'readComputedStyle'> = {},
): Promise<IsolatedRasterBase | undefined> {
	const doc = stage.ownerDocument as Document & FontStyleDocumentLike & LinkStyleDocumentLike;
	let unmark = () => {};
	try {
		unmark = markDeclaredVisibility(stage);
		for (const [target, tag] of targets) target.setAttribute(RASTER_TARGET_ATTRIBUTE, tag);
		const prepared = await settleWithin(
			prepareForeignObjectClone(stage, doc, {
				untransformedRoot: true,
				...(options.readComputedStyle ? { readComputedStyle: options.readComputedStyle } : {}),
				...(options.externalFonts ? { externalFonts: options.externalFonts() } : {}),
			}),
			RASTER_TIMEOUT_MS,
			undefined,
		);
		if (!prepared) return undefined;
		return {
			...buildIsolatedTemplate(prepared.clone, RASTER_TARGET_ATTRIBUTE),
			fontFaceCss: prepared.fontFaceCss,
			allEmbedded: prepared.allEmbedded,
		};
	} finally {
		unmark();
		for (const target of targets.keys()) target.removeAttribute(RASTER_TARGET_ATTRIBUTE);
	}
}

const liveStageOnly = () => Promise.reject(new Error('Drawn from the live stage instead.'));

/**
 * Draw every request (see the module comment), in order. An entry is `undefined` when no
 * strategy could produce readable pixels for it.
 */
export async function rasterizeIsolatedBatch(
	stage: HTMLElement,
	requests: readonly IsolatedRasterRequest[],
	naturalWidth: number,
	naturalHeight: number,
	options: IsolatedRasterBatchOptions,
): Promise<(PdfImageResource | undefined)[]> {
	const doc = stage.ownerDocument;
	const tags = new Map<Element, string>();
	for (const { target } of requests) if (!tags.has(target)) tags.set(target, String(tags.size));
	const base = requests.length
		? await prepareIsolatedRasterBase(stage, tags, options).catch(() => undefined)
		: undefined;
	const one = (request: IsolatedRasterRequest) =>
		rasterizeIsolated(
			stage,
			request.target,
			request.mode,
			request.rect,
			naturalWidth,
			naturalHeight,
			options,
		);
	const out: (PdfImageResource | undefined)[] = [];
	for (const request of requests) {
		const target = base?.allEmbedded ? base.targets.get(tags.get(request.target)!) : undefined;
		if (!base || target === undefined) {
			out.push(await one(request));
			continue;
		}
		const { rect, mode } = request;
		const output = {
			width: Math.max(1, Math.round(rect.width * options.scale)),
			height: Math.max(1, Math.round(rect.height * options.scale)),
		};
		let shadowStyle: string | undefined;
		try {
			if (mode === 'shadow')
				shadowStyle = shadowTargetStyle(stage, request.target, options.readComputedStyle);
		} catch {
			out.push(await one(request));
			continue;
		}
		const work = (async () => {
			const markup = isolatedMarkup(base.template, target, mode, {
				...(shadowStyle === undefined ? {} : { shadowStyle }),
			});
			const body = {
				bodyMarkup: foreignObjectBodyMarkup(markup, base.fontFaceCss, {
					width: naturalWidth,
					height: naturalHeight,
				}),
				allEmbedded: true,
				naturalWidth,
				naturalHeight,
			};
			const { canvas } = await rasterizeWindow(body, rect, output, undefined, {
				html2canvasFallback: liveStageOnly,
			});
			return await drawableToPdfImage(doc, canvas, canvas.width, canvas.height, 8192);
		})();
		// A draw this path could not make goes the one-at-a-time way, html2canvas fallback included.
		out.push((await settleWithin(work, RASTER_TIMEOUT_MS, undefined)) ?? (await one(request)));
	}
	return out;
}

/**
 * The document's external font CSS for {@link IsolatedRasterOptions.externalFonts}, fetched once
 * and again only when its font stylesheets change (a later slide added one).
 */
export function externalFontCache(doc: Document): () => Promise<ExternalFontFaceCssResult> {
	let key: string | undefined;
	let value: Promise<ExternalFontFaceCssResult> | undefined;
	return () => {
		const links = Array.from(doc.querySelectorAll('link[rel="stylesheet"]'), (link) =>
			link.getAttribute('href'),
		).join('\n');
		if (value === undefined || links !== key) {
			key = links;
			value = collectExternalFontFaceCss(doc, fetchAsDataUrl);
		}
		return value;
	};
}
