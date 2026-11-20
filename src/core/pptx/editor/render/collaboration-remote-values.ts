/**
 * collaboration-remote-values.ts: what a collaboration peer may put into the deck.
 *
 * Everything read back from the shared Y.Doc was written by some peer, and a peer can write
 * anything the CRDT can carry. Before a slide read from the document reaches a binding (and
 * through it the DOM, a download or a new browser tab), every element id is put in the safe
 * token form the loader builds ids in, and every resource URL and MIME type is checked
 * against the rule of its field. A value that fails is left out and reported as a
 * `COLLABORATION_REMOTE_VALUE_REJECTED` warning on its slide, which the bindings show.
 *
 * Those warnings describe this client's view only, so they never travel back into the
 * shared document (see {@link sharedSlideRecord}).
 */

import type { PptxCompatibilityWarning, PptxElement, PptxSlide } from 'ooxml-core/pptx';

import { safeIdToken } from '../../../opc/part-name';
import {
	isAllowedResourceUrl,
	isWellFormedContentType,
	normalizeMimeType,
	parseDataUrlHeader,
	type ResourceRule,
} from '../../../opc/safe-resource-url';
import type { CollaborationMedia } from './collaboration-media';

/** Warning code for a value from a peer that was left out. */
export const REMOTE_VALUE_REJECTED = 'COLLABORATION_REMOTE_VALUE_REJECTED';

/** Warning code for a picture or media file a peer could not share because it is too large. */
export const REMOTE_MEDIA_WITHHELD = 'COLLABORATION_MEDIA_WITHHELD';

/** One value a peer sent that was left out of this client's deck. */
export interface RemoteValueRejection {
	/** The element it belonged to; absent when the element itself was left out. */
	readonly elementId?: string | undefined;
	readonly field: string;
	/** `withheld`: the file was too large to share; otherwise the value was not allowed. */
	readonly reason?: 'invalid' | 'withheld' | undefined;
}

const VIEW_ONLY_CODES: ReadonlySet<string> = new Set([
	REMOTE_VALUE_REJECTED,
	REMOTE_MEDIA_WITHHELD,
]);

/**
 * Element fields that hold a resource URL, with where each may point. Web URLs are the
 * linked pictures and media a loader keeps when the host allows external content; paths
 * are package parts the loader did not decode.
 */
const ELEMENT_RESOURCES: Readonly<Record<string, ResourceRule>> = {
	imageData: { kind: 'image', web: true },
	svgData: { kind: 'image' },
	posterImage: { kind: 'image', web: true, path: true },
	posterFrameData: { kind: 'image', web: true },
	previewImage: { kind: 'image', path: true },
	previewImageData: { kind: 'image' },
	mediaData: { kind: 'media', web: true },
	modelData: { kind: 'model' },
	oleEmbeddedData: { kind: 'file' },
};

const SLIDE_RESOURCES: Readonly<Record<string, ResourceRule>> = {
	backgroundImage: { kind: 'image', web: true },
};

const CONTENT_TYPE_FIELDS = [
	'mediaMimeType',
	'modelMimeType',
	'mediaReferenceContentType',
	'oleEmbeddedMimeType',
] as const;

const isRecord = (value: unknown): value is Record<string, unknown> =>
	typeof value === 'object' && value !== null && !Array.isArray(value);

function dropInvalidResources(
	record: Record<string, unknown>,
	rules: Readonly<Record<string, ResourceRule>>,
	rejected: RemoteValueRejection[],
	elementId: string | undefined,
	media: CollaborationMedia | undefined,
): void {
	for (const [field, rule] of Object.entries(rules)) {
		const value = record[field];
		// A URL this page made for the room's bytes is checked against the type they were sent as.
		if (
			value !== undefined &&
			!isAllowedResourceUrl(value, rule) &&
			media?.accepts(value, rule.kind) !== true
		) {
			delete record[field];
			rejected.push(elementId === undefined ? { field } : { elementId, field });
		}
	}
}

/**
 * Check one element read from the shared document, in place, including the children of a
 * group. Returns false when the element has no usable id and must be left out.
 */
export function sanitizeRemoteElement(
	element: Record<string, unknown>,
	rejected: RemoteValueRejection[],
	media?: CollaborationMedia,
): boolean {
	if (typeof element.id !== 'string' || element.id.length === 0) {
		rejected.push({ field: 'id' });
		return false;
	}
	const elementId = safeIdToken(element.id);
	element.id = elementId;
	dropInvalidResources(element, ELEMENT_RESOURCES, rejected, elementId, media);
	for (const field of CONTENT_TYPE_FIELDS) {
		if (element[field] !== undefined && !isWellFormedContentType(element[field])) {
			delete element[field];
			rejected.push({ elementId, field });
		}
	}
	// A browser types an embedded file by its URL, so a separate type must agree with it.
	const embedded = parseDataUrlHeader(element.oleEmbeddedData);
	if (
		embedded &&
		element.oleEmbeddedMimeType !== undefined &&
		normalizeMimeType(String(element.oleEmbeddedMimeType).split(';', 1)[0]) !== embedded.mime
	) {
		delete element.oleEmbeddedMimeType;
		rejected.push({ elementId, field: 'oleEmbeddedMimeType' });
	}
	if (element.children !== undefined) {
		element.children = Array.isArray(element.children)
			? sanitizeRemoteElements(element.children, rejected, media)
			: [];
	}
	return true;
}

/** {@link sanitizeRemoteElement} over a list, leaving out the elements it refuses. */
export function sanitizeRemoteElements(
	elements: readonly unknown[],
	rejected: RemoteValueRejection[],
	media?: CollaborationMedia,
): PptxElement[] {
	return elements.filter(
		(element) => isRecord(element) && sanitizeRemoteElement(element, rejected, media),
	) as PptxElement[];
}

/**
 * Check the slide-level values of a slide read from the shared document, in place, and record
 * every rejection on it (its own and its elements') as a warning.
 */
export function sanitizeRemoteSlide(
	slide: Record<string, unknown>,
	rejected: RemoteValueRejection[],
	media?: CollaborationMedia,
): void {
	dropInvalidResources(slide, SLIDE_RESOURCES, rejected, undefined, media);
	if (Array.isArray(slide.notesShapes)) {
		slide.notesShapes = sanitizeRemoteElements(slide.notesShapes, rejected, media);
	}
	if (rejected.length === 0) {
		return;
	}
	const slideId = typeof slide.id === 'string' ? slide.id : undefined;
	const existing = Array.isArray(slide.warnings)
		? (slide.warnings as PptxCompatibilityWarning[])
		: [];
	slide.warnings = [
		...existing,
		...rejected.map((rejection): PptxCompatibilityWarning => ({
			code: rejection.reason === 'withheld' ? REMOTE_MEDIA_WITHHELD : REMOTE_VALUE_REJECTED,
			message:
				rejection.reason === 'withheld'
					? `A collaborator's file for ${rejection.field} is too large to share, so it is not shown.`
					: `A collaborator sent a value for ${rejection.field} that is not allowed, so it was left out.`,
			severity: 'warning',
			scope: rejection.elementId === undefined ? 'slide' : 'element',
			...(slideId === undefined ? {} : { slideId }),
			...(rejection.elementId === undefined ? {} : { elementId: rejection.elementId }),
		})),
	];
}

/**
 * The slide as it should be written to the shared document: without the rejection warnings
 * this client recorded, which describe its own view and are not part of the deck.
 */
export function sharedSlideRecord(slide: PptxSlide): Record<string, unknown> {
	const record = slide as unknown as Record<string, unknown>;
	if (!slide.warnings?.some((warning) => VIEW_ONLY_CODES.has(warning.code))) {
		return record;
	}
	const warnings = slide.warnings.filter((warning) => !VIEW_ONLY_CODES.has(warning.code));
	return { ...record, warnings: warnings.length > 0 ? warnings : undefined };
}
