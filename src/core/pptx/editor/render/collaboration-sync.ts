/**
 * collaboration-sync.ts: Framework-agnostic CRDT sync utilities for the
 * pptx-viewer collaboration stack (Yjs backend).
 *
 * Exports:
 *  - Structural Yjs interfaces (no hard yjs import - bindings pass live instances)
 *  - YjsFactories: factory interface bindings implement using `new Y.Map()` etc.
 *  - writeElementToYMap / readElementFromYMap: PptxElement <-> YMapLike
 *  - writeSlideToYMap / readSlideFromYMap: PptxSlide <-> YMapLike
 *  - writeSlidesToYDoc / readSlidesFromYDoc: PptxSlide[] <-> Y.Doc
 *  - observeYDocSlides: register a change listener on the pptx:slides array
 *  - re-exports of the text codec (collaboration-text-codec.ts)
 *
 * Y.Doc schema:
 *   pptx:slides  - Y.Array of slide Y.Maps
 *   Each slide Y.Map has scalar keys + `_`-prefixed JSON blobs + `elements`
 *   Each element Y.Map has scalar keys + `_`-prefixed JSON blobs + `textBody`
 *   textBody is a Y.Text with one delta-op per TextSegment
 *
 * NOTE: packages/tools' pptx-codec.ts implements a similar schema but with
 * different complex-field key prefixes (e.g. `_textStyle` vs `_ts` here); the
 * two doc layouts are NOT interchangeable on the same Y.Doc.
 *
 * Prefer `reconcileSlidesInYDoc` (collaboration-reconcile.ts) over
 * `writeSlidesToYDoc` for live editing: it updates only what changed instead
 * of replacing the whole slides array, so concurrent edits merge per
 * slide/element/field rather than colliding at document granularity.
 */

import type { PptxSlide, PptxElement, TextSegment } from 'ooxml-core/pptx';

import {
	ASSET_ELEMENT_FIELDS,
	getAssetsMap,
	isAssetRefKey,
	isAssetVersionKey,
	readAssetFields,
	writeAssetFields,
} from './collaboration-assets';
import { type CollaborationMedia, collaborationMedia } from './collaboration-media';
import { collectSharedMedia, sharedValue, stringifyShared } from './collaboration-media-token';
import {
	inboundString,
	mapInbound,
	type MediaValueIssue,
	parseInbound,
} from './collaboration-media-values';
import {
	type RemoteValueRejection,
	sanitizeRemoteElements,
	sanitizeRemoteSlide,
	sharedSlideRecord,
} from './collaboration-remote-values';
import type { YTextLike } from './collaboration-text-codec';
import { encodeTextBody, decodeTextBody, isYTextLike } from './collaboration-text-codec';
import type { YTextEditableLike } from './collaboration-text-merge';
import type { TrackedTextPosition } from './collaboration-text-positions';
import type { TextSessionPositions } from './collaboration-text-session';
import { inlineListBodyText } from './inline-list-body';

export * from './collaboration-assets';
export * from './collaboration-media';
export * from './collaboration-remote-values';
export * from './collaboration-text-codec';

// ---------------------------------------------------------------------------
// Structural Yjs interfaces (no 'yjs' import; bindings supply live instances)
// ---------------------------------------------------------------------------

export interface YMapLike {
	get: (key: string) => unknown;
	set: (key: string, value: unknown) => void;
	delete: (key: string) => void;
	forEach: (cb: (value: unknown, key: string) => void) => void;
}

/** Shape of the Yjs transaction passed to (deep) observers. */
export interface YTransactionLike {
	origin?: unknown;
}

export type YDeepObserver = (events?: unknown, transaction?: YTransactionLike) => void;

export interface YArrayLike {
	readonly length: number;
	get: (index: number) => unknown;
	push: (items: unknown[]) => void;
	delete: (index: number, length?: number) => void;
	insert: (index: number, items: unknown[]) => void;
	toArray: () => unknown[];
	observe: (handler: () => void) => void;
	unobserve: (handler: () => void) => void;
	observeDeep: (handler: YDeepObserver) => void;
	unobserveDeep: (handler: YDeepObserver) => void;
}

export interface YDocLike {
	getMap: (name: string) => YMapLike;
	getArray: (name: string) => YArrayLike;
	transact: (fn: () => void, origin?: unknown) => void;
}

export interface YjsFactories {
	createMap: () => YMapLike;
	createArray: () => YArrayLike;
	createText: () => YTextLike;
	/** Track character identities using the binding's existing Yjs runtime. */
	createTextPositions?:
		| ((text: YTextEditableLike) => TextSessionPositions<TrackedTextPosition>)
		| undefined;
}

// ---------------------------------------------------------------------------
// Y.Doc schema constants
// ---------------------------------------------------------------------------

export const YDOC_SLIDES_KEY = 'pptx:slides';
export const YDOC_META_KEY = 'pptx:meta';

export const SCALAR_ELEMENT_KEYS: ReadonlySet<string> = new Set([
	'id',
	'type',
	'x',
	'y',
	'width',
	'height',
	'xEmu',
	'yEmu',
	'widthEmu',
	'heightEmu',
	'rotation',
	'shapeId',
	'relationshipSourcePath',
	'skewX',
	'skewY',
	'flipHorizontal',
	'flipVertical',
	'hidden',
	'opacity',
	'text',
	'name',
	'altText',
	'title',
	'isDecorative',
	'preferRelativeResize',
	'oleUpdateAutomatic',
	'shapeType',
	'imagePath',
	'imageData',
	'svgData',
	'svgPath',
	'cropLeft',
	'cropTop',
	'cropRight',
	'cropBottom',
	'fillRectLeft',
	'fillRectTop',
	'fillRectRight',
	'fillRectBottom',
	'tileOffsetX',
	'tileOffsetY',
	'tileScaleX',
	'tileScaleY',
	'tileFlip',
	'tileAlignment',
	'dpi',
	'pathData',
	'pathWidth',
	'pathHeight',
	'mediaType',
	'mediaPath',
	'mediaMimeType',
	'mediaReferenceKind',
	'mediaReferenceName',
	'mediaReferenceContentType',
	'trimStartMs',
	'trimEndMs',
	'posterFramePath',
	'fullScreen',
	'loop',
	'fadeInDuration',
	'fadeOutDuration',
	'volume',
	'autoPlay',
	'playAcrossSlides',
	'hideWhenNotPlaying',
	'playbackSpeed',
	'mediaMissing',
	'isLinked',
	'oleTarget',
	'oleProgId',
	'oleName',
	'oleClsId',
	'oleObjectType',
	'oleFileExtension',
	'fileName',
	'externalPath',
	'previewImage',
	'oleShowAsIcon',
	'oleImgW',
	'oleImgH',
	'oleEmbeddedFileName',
	'oleEmbeddedMimeType',
	'oleEmbeddedByteSize',
	'oleFollowColorScheme',
	// Set by in-place OLE content editing until the next save re-embeds the
	// payload (`ole-edit-commit.ts` in `ooxml-core/pptx`); a plain flag.
	'oleContentDirty',
	// GroupPptxElement: exact EMU for the group's own `a:chOff`/`a:chExt`
	// (child coordinate space). See `group-xfrm-preservation.ts` in
	// `ooxml-core/pptx`.
	'chOffXEmu',
	'chOffYEmu',
	'chExtWidthEmu',
	'chExtHeightEmu',
	'inkPaths',
	'inkColors',
	'inkWidths',
	'inkOpacities',
	'inkTool',
	'inkPartPath',
	'zoomType',
	'zoomImageType',
	'zoomShowBackground',
	'returnToParent',
	'transitionDurationMs',
	'targetSlideIndex',
	'targetSectionId',
	'summaryLayout',
	'modelPath',
	'modelMimeType',
	'posterImage',
	'linkedTxbxId',
	'linkedTxbxSeq',
	'promptText',
	'placeholderType',
	'placeholderSz',
	'placeholderOrient',
	// Resolved from the slide master rather than authored, but it travels with
	// `text` on purpose: the two are compared by the save writer to tell an
	// inherited footer from one edited on this slide. See
	// `collaboration-field-schema.ts`.
	'inheritedPlaceholderText',
]);

export const COMPLEX_ELEMENT_FIELDS: Readonly<Record<string, string>> = {
	textStyle: '_ts',
	shapeStyle: '_ss',
	shapeAdjustments: '_sa',
	adjustmentHandles: '_ah',
	tableData: '_td',
	chartData: '_cd',
	smartArtData: '_smad',
	children: '_ch',
	paragraphIndents: '_pi',
	rawXml: '_rx',
	inheritedTransform: '_it',
	extLstXml: '_elx',
	actionClick: '_ac',
	actionHover: '_av',
	locks: '_lk',
	imageEffects: '_ie',
	cropShape: '_cr',
	bookmarks: '_mb',
	captionTracks: '_ct',
	audioCdStart: '_acd1',
	audioCdEnd: '_acd2',
	rawMediaReferenceXml: '_mrx',
	metadata: '_md',
	groupFill: '_gf',
	groupEffectStyle: '_ges',
	inkPointPressures: '_ipp',
	inkPointTiltX: '_iptx',
	inkPointTiltY: '_ipty',
	inkStrokes: '_cis',
	inkPartRawXml: '_cirx',
	summaryTargets: '_zst',
	extensionXml: '_ext',
	customGeometryPaths: '_cgp',
	customGeometryRawData: '_cgr',
	customGeometryAdjustHandlesXY: '_cgx',
	customGeometryAdjustHandlesPolar: '_cgo',
	customGeometryConnectionSites: '_cgc',
	customGeometryTextRect: '_cgt',
	// A 3D model's authored camera, transform and lights (read-only view of `rawXml`).
	scene: '_m3s',
};
const REV_COMPLEX_ELEMENT: Record<string, string> = Object.fromEntries(
	Object.entries(COMPLEX_ELEMENT_FIELDS).map(([k, v]) => [v, k]),
);

export const SCALAR_SLIDE_KEYS: ReadonlySet<string> = new Set([
	'id',
	'rId',
	// The real `p:sldIdLst` id (256+ on a PowerPoint deck). Sections reference
	// it, so peers must agree on it exactly as they do on `rId`.
	'slideId',
	'sourceSlideId',
	'name',
	'layoutPath',
	'layoutName',
	'slideNumber',
	'hidden',
	'sectionName',
	'sectionId',
	'backgroundColor',
	'backgroundImage',
	'backgroundGradient',
	'backgroundShadeToTitle',
	'notes',
	'notesCSldName',
	'backgroundShowAnimation',
	'showMasterShapes',
	'showMasterPhAnim',
	'isDirty',
]);

export const COMPLEX_SLIDE_FIELDS: Readonly<Record<string, string>> = {
	transition: '_tr',
	animations: '_an',
	animationTimelineAnchors: '_ata',
	nativeAnimations: '_na',
	rawTiming: '_rt',
	notesSegments: '_ns',
	notesShapes: '_nsh',
	notesClrMapOverride: '_ncm',
	comments: '_cm',
	warnings: '_wa',
	rawXml: '_rx',
	clrMapOverride: '_cm2',
	guides: '_gu',
	customerData: '_cu',
	activeXControls: '_ax',
	legacyVmlElements: '_lvml',
	backgroundPattern: '_bp',
	backgroundGradientFill: '_bgf',
	backgroundImageProperties: '_bip',
	backgroundStyleRef: '_bsr',
	modernCommentPart: '_mc',
	headerFooterFlags: '_hff',
	slideSynchronization: '_sync',
};
const REV_COMPLEX_SLIDE: Record<string, string> = Object.fromEntries(
	Object.entries(COMPLEX_SLIDE_FIELDS).map(([k, v]) => [v, k]),
);

// ---------------------------------------------------------------------------
// Element serialization
// ---------------------------------------------------------------------------

/**
 * The element as it is written to the room. An embedded audio or video element names only
 * its part; the URL this page plays it from is shared as its `mediaData`, so a participant
 * without the deck's file can play it too. Page-local URLs are written as media tokens by
 * the writers (see collaboration-media-token.ts).
 */
export function sharedElementRecord(
	element: PptxElement,
	media: CollaborationMedia | undefined,
): Record<string, unknown> {
	const rec = element as unknown as Record<string, unknown>;
	if (rec.type !== 'media' || rec.mediaData !== undefined || typeof rec.mediaPath !== 'string') {
		return rec;
	}
	const url = media?.localMediaUrl(rec.mediaPath);
	return url === undefined ? rec : { ...rec, mediaData: url };
}

export function writeElementToYMap(
	element: PptxElement,
	ymap: YMapLike,
	factories: YjsFactories,
	assets: YMapLike,
	media?: CollaborationMedia,
): void {
	const rec = sharedElementRecord(element, media);
	for (const [key, value] of Object.entries(rec)) {
		if (value === undefined || ASSET_ELEMENT_FIELDS.has(key)) {
			continue;
		}
		if (SCALAR_ELEMENT_KEYS.has(key)) {
			ymap.set(key, sharedValue(value));
		} else if (key === 'textSegments') {
			if (Array.isArray(value)) {
				const ytext = factories.createText();
				encodeTextBody(value, ytext);
				ymap.set('textBody', ytext);
			}
		} else if (COMPLEX_ELEMENT_FIELDS[key]) {
			ymap.set(COMPLEX_ELEMENT_FIELDS[key], stringifyShared(value));
		}
	}
	writeAssetFields(rec.id as string, rec, ymap, assets);
}

/** Pass the top-level string fields of `record` through {@link inboundString}. */
function inboundFields(
	record: Record<string, unknown>,
	fields: Iterable<string>,
	media: CollaborationMedia,
	issues: MediaValueIssue[],
): void {
	for (const field of fields) {
		const value = record[field];
		if (typeof value === 'string') {
			const next = inboundString(media, field, value, issues);
			if (next === undefined) {
				delete record[field];
			} else {
				record[field] = next;
			}
		}
	}
}

/**
 * Read one element. With `media`, the room's media tokens become this page's URLs and the
 * values this viewer must not show are left out; `issues` collects the ones to report.
 */
export function readElementFromYMap(
	ymap: YMapLike,
	assets: YMapLike,
	media?: CollaborationMedia,
	issues: MediaValueIssue[] = [],
): PptxElement {
	const element: Record<string, unknown> = {};
	ymap.forEach((value: unknown, key: string) => {
		if (key === 'textBody') {
			if (isYTextLike(value)) {
				const segments = decodeTextBody(value);
				element.textSegments = media ? mapInbound(media, segments, issues) : segments;
			}
		} else if (isAssetRefKey(key) || isAssetVersionKey(key)) {
			// Ref pointers are resolved by readAssetFields below; version
			// counters are an internal sync token, never a PptxElement field.
		} else if (REV_COMPLEX_ELEMENT[key]) {
			try {
				element[REV_COMPLEX_ELEMENT[key]] = parseInbound(media, value as string, issues);
			} catch {
				/* skip */
			}
		} else {
			element[key] = value;
		}
	});
	readAssetFields(ymap, assets, element);
	if (media) {
		inboundFields(element, Object.keys(element), media, issues);
	}
	// The scalar mirror is last-writer-wins; rich text merges character-wise.
	// Derive its body after a remote merge, but retain legacy scalar-only text
	// when a document has an empty textSegments array and no encoded carriers.
	if (Array.isArray(element.textSegments) && element.textSegments.length > 0) {
		element.text = inlineListBodyText(element.textSegments as TextSegment[]);
	}
	return element as unknown as PptxElement;
}

// ---------------------------------------------------------------------------
// Slide serialization
// ---------------------------------------------------------------------------

export function writeSlideToYMap(
	slide: PptxSlide,
	ymap: YMapLike,
	factories: YjsFactories,
	assets: YMapLike,
	media?: CollaborationMedia,
): void {
	const rec = sharedSlideRecord(slide);
	for (const key of SCALAR_SLIDE_KEYS) {
		if (rec[key] !== undefined) {
			ymap.set(key, sharedValue(rec[key]));
		}
	}
	for (const [original, prefixed] of Object.entries(COMPLEX_SLIDE_FIELDS)) {
		if (rec[original] !== undefined) {
			ymap.set(prefixed, stringifyShared(rec[original]));
		}
	}
	const elemArr = factories.createArray();
	for (const el of slide.elements) {
		const elemMap = factories.createMap();
		writeElementToYMap(el, elemMap, factories, assets, media);
		elemArr.push([elemMap]);
	}
	ymap.set('elements', elemArr);
}

/**
 * Read one slide written by any peer. Its elements and resource fields are checked on the
 * way out (see collaboration-remote-values.ts), so what is returned is safe to render.
 */
export function readSlideFromYMap(
	ymap: YMapLike,
	assets: YMapLike,
	media?: CollaborationMedia,
): PptxSlide {
	const slide: Record<string, unknown> = {};
	const slideIssues: MediaValueIssue[] = [];
	for (const key of SCALAR_SLIDE_KEYS) {
		const v = ymap.get(key);
		if (v !== undefined) {
			slide[key] = v;
		}
	}
	for (const [prefixed, original] of Object.entries(REV_COMPLEX_SLIDE)) {
		const v = ymap.get(prefixed) as string | undefined;
		if (v !== undefined) {
			try {
				slide[original] = parseInbound(media, v, slideIssues);
			} catch {
				/* skip */
			}
		}
	}
	if (media) {
		inboundFields(slide, SCALAR_SLIDE_KEYS, media, slideIssues);
	}
	const rejected: RemoteValueRejection[] = slideIssues.map(({ field, reason }) => ({
		field,
		reason,
	}));
	const elemArr = ymap.get('elements') as YArrayLike | undefined;
	const elements: PptxElement[] = [];
	if (elemArr) {
		for (let i = 0; i < elemArr.length; i++) {
			const issues: MediaValueIssue[] = [];
			const element = readElementFromYMap(elemArr.get(i) as YMapLike, assets, media, issues);
			const id = (element as { id?: unknown }).id;
			for (const { field, reason } of issues) {
				rejected.push(
					typeof id === 'string' ? { elementId: id, field, reason } : { field, reason },
				);
			}
			elements.push(element);
		}
	}
	slide.elements = sanitizeRemoteElements(elements, rejected, media);
	sanitizeRemoteSlide(slide, rejected, media);
	return slide as unknown as PptxSlide;
}

// ---------------------------------------------------------------------------
// Y.Doc-level helpers
// ---------------------------------------------------------------------------

/**
 * Replace the full slides array in the Y.Doc. Coarse: prefer
 * `reconcileSlidesInYDoc` for live editing; this remains suitable for
 * one-shot seeding of an empty document.
 */
export function writeSlidesToYDoc(
	slides: PptxSlide[],
	ydoc: YDocLike,
	factories: YjsFactories,
	origin?: unknown,
): void {
	const assets = getAssetsMap(ydoc);
	const media = collaborationMedia(ydoc);
	const shared = collectSharedMedia(() =>
		ydoc.transact(() => {
			const arr = ydoc.getArray(YDOC_SLIDES_KEY);
			if (arr.length > 0) {
				arr.delete(0, arr.length);
			}
			for (const slide of slides) {
				const ymap = factories.createMap();
				writeSlideToYMap(slide, ymap, factories, assets, media);
				arr.push([ymap]);
			}
		}, origin),
	);
	void media.upload(shared, origin);
}

/**
 * Read the deck from the document. A slide id that appears more than once (two
 * peers seeded the same empty room concurrently) is read once, from its first
 * copy, which is also the copy `reconcileSlidesInYDoc` keeps.
 */
export function readSlidesFromYDoc(ydoc: YDocLike): PptxSlide[] {
	const assets = getAssetsMap(ydoc);
	const media = collaborationMedia(ydoc);
	const arr = ydoc.getArray(YDOC_SLIDES_KEY);
	const slides: PptxSlide[] = [];
	const seen = new Set<string>();
	for (let i = 0; i < arr.length; i++) {
		const map = arr.get(i) as YMapLike;
		const id = typeof map?.get === 'function' ? map.get('id') : undefined;
		if (typeof id === 'string') {
			if (seen.has(id)) {
				continue;
			}
			seen.add(id);
		}
		slides.push(readSlideFromYMap(map, assets, media));
	}
	return slides;
}

/**
 * Observe (deeply) the pptx:slides array, and the room's media (the bytes of a picture
 * arrive after the slide that shows it) and this viewer's media settings. The handler
 * receives the Yjs events plus the transaction, so callers can skip their own writes by
 * checking `transaction.origin` (see LOCAL_SYNC_ORIGIN in collaboration-reconcile.ts).
 */
export function observeYDocSlides(ydoc: YDocLike, onChange: YDeepObserver): () => void {
	const arr = ydoc.getArray(YDOC_SLIDES_KEY);
	arr.observeDeep(onChange);
	const unobserveMedia = collaborationMedia(ydoc).observe(onChange);
	return () => {
		arr.unobserveDeep(onChange);
		unobserveMedia();
	};
}
