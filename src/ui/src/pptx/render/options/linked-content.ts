/**
 * Linked web content in a viewer's own view of the deck.
 *
 * A deck can link a picture, fill, background, bullet, table texture, poster,
 * audio or video to a web address instead of embedding it. The document model
 * keeps that address as the deck's content: it is saved, collaborated and
 * undone like any other field. Whether the browser may fetch it is a choice
 * each viewer makes for itself (`external-content-policy`), so the gate sits
 * where the model becomes markup: every surface that draws or plays the deck
 * reads it through {@link linkedContentView}, which withholds the addresses
 * until that viewer allows them. Enabling the content redraws; nothing
 * reloads and nothing in the model changes.
 *
 * @module render/options/linked-content
 */

/**
 * Model fields whose value a surface hands the browser to fetch (an image
 * `src`, a CSS `url()`, a media source, or the address an online video embed
 * is derived from). Hyperlinks and other addresses that open only when
 * clicked are not listed: they are never fetched by drawing the deck.
 */
const FETCHED_FIELDS: ReadonlySet<string> = new Set([
	'imageData',
	'imagePath',
	'svgData',
	'svgPath',
	'posterImage',
	'posterFrameData',
	'posterFramePath',
	'previewImageData',
	'mediaData',
	'mediaPath',
	'fillImageUrl',
	'backgroundImage',
	'imageDataUrl',
	'textFillBlipUrl',
	'backgroundImageFillPath',
	'backgroundImageFillData',
]);

/** A table style's picture fill (`ParsedTableStyleFill.image`) names its picture with these. */
const TABLE_STYLE_IMAGE_KEY = 'image';
const TABLE_STYLE_IMAGE_FIELDS: ReadonlySet<string> = new Set(['path', 'data']);

/** Whether `value` is a web address a browser would fetch (`http:` or `https:`). */
export function isWebAddress(value: unknown): value is string {
	return typeof value === 'string' && /^https?:\/\//iu.test(value);
}

function isFetchedField(key: string, parentKey: string | undefined): boolean {
	return (
		FETCHED_FIELDS.has(key) ||
		(parentKey === TABLE_STYLE_IMAGE_KEY && TABLE_STYLE_IMAGE_FIELDS.has(key))
	);
}

function isTraversable(value: unknown): value is Record<string, unknown> | unknown[] {
	if (Array.isArray(value)) {
		return true;
	}
	if (typeof value !== 'object' || value === null) {
		return false;
	}
	const proto = Object.getPrototypeOf(value) as unknown;
	return proto === Object.prototype || proto === null;
}

// Model objects are replaced, never mutated, so a result per object stays valid.
const withheldByValue = new WeakMap<object, unknown>();
const linksByValue = new WeakMap<object, boolean>();
// Each withheld copy back to the model object it was made from (`restoreLinkedContent`).
const modelByView = new WeakMap<object, object>();

function withhold(value: unknown, parentKey: string | undefined, seen: Set<object>): unknown {
	if (!isTraversable(value) || seen.has(value)) {
		return value;
	}
	const memo = withheldByValue.get(value);
	if (memo !== undefined) {
		return memo;
	}
	seen.add(value);
	let result: unknown = value;
	if (Array.isArray(value)) {
		let copy: unknown[] | undefined;
		value.forEach((item, index) => {
			const next = withhold(item, parentKey, seen);
			if (next !== item) {
				copy ??= [...value];
				copy[index] = next;
			}
		});
		result = copy ?? value;
	} else {
		let copy: Record<string, unknown> | undefined;
		for (const [key, item] of Object.entries(value)) {
			if (isFetchedField(key, parentKey) && isWebAddress(item)) {
				copy ??= { ...value };
				delete copy[key];
				continue;
			}
			const next = withhold(item, key, seen);
			if (next !== item) {
				copy ??= { ...value };
				copy[key] = next;
			}
		}
		result = copy ?? value;
	}
	seen.delete(value);
	withheldByValue.set(value, result);
	if (result !== value) {
		modelByView.set(result as object, value);
	}
	return result;
}

function restore(value: unknown, seen: Set<object>): unknown {
	if (!isTraversable(value) || seen.has(value)) {
		return value;
	}
	const model = modelByView.get(value);
	if (model) {
		return model;
	}
	seen.add(value);
	let result: unknown = value;
	if (Array.isArray(value)) {
		const next = value.map((item) => restore(item, seen));
		result = next.some((item, index) => item !== value[index]) ? next : value;
	} else {
		let copy: Record<string, unknown> | undefined;
		for (const [key, item] of Object.entries(value)) {
			const next = restore(item, seen);
			if (next !== item) {
				copy ??= { ...value };
				copy[key] = next;
			}
		}
		result = copy ?? value;
	}
	seen.delete(value);
	return result;
}

function links(value: unknown, parentKey: string | undefined, seen: Set<object>): boolean {
	if (!isTraversable(value) || seen.has(value)) {
		return false;
	}
	const memo = linksByValue.get(value);
	if (memo !== undefined) {
		return memo;
	}
	seen.add(value);
	const found = Array.isArray(value)
		? value.some((item) => links(item, parentKey, seen))
		: Object.entries(value).some(([key, item]) =>
				isFetchedField(key, parentKey) && isWebAddress(item) ? true : links(item, key, seen),
			);
	seen.delete(value);
	linksByValue.set(value, found);
	return found;
}

/**
 * `value` (a slide, an element, a list of them, a table style map, ...) with
 * every linked web address a surface would fetch removed. Objects without one
 * are returned as they are, so a surface keeps its identity checks and memos;
 * the rest are shallow copies. Pure: the model itself is never changed.
 */
export function withholdLinkedContent<T>(value: T): T {
	return withhold(value, undefined, new Set()) as T;
}

/**
 * An edit a surface built from what it drew, turned back into model data: every part the edit
 * kept from a {@link withholdLinkedContent} copy is replaced by the model object that copy was
 * made from, links included. Parts the edit rebuilt keep what the edit made. For a surface that
 * commits a whole structure (SmartArt data, chart data) derived from the element it drew.
 */
export function restoreLinkedContent<T>(value: T): T {
	return restore(value, new Set()) as T;
}

/** Whether `value` holds a linked web address that {@link withholdLinkedContent} would remove. */
export function hasLinkedContent(value: unknown): boolean {
	return links(value, undefined, new Set());
}

/**
 * What a surface draws: `value` itself once this viewer allows linked content
 * (`resolveExternalContentView`), otherwise {@link withholdLinkedContent}.
 */
export function linkedContentView<T>(value: T, allowed: boolean): T {
	return allowed ? value : withholdLinkedContent(value);
}

/**
 * One address as a surface may hand it to the browser: `undefined` for a web address this
 * viewer has not allowed. For a preview that draws a single field (an inspector thumbnail)
 * rather than a whole slide or element.
 */
export function linkedContentUrl(url: string | undefined, allowed: boolean): string | undefined {
	return allowed || !isWebAddress(url) ? url : undefined;
}
