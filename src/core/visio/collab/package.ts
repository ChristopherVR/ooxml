// Package-level helpers of the Visio mapping: relationships, part paths, the parts that stay local,
// base64 and a cheap fingerprint.
import { sha256 } from '../../digest/sha256';
import { buildXml, parseXml } from '../../xml/index';
import { attrOf } from './scan';

export interface Relationship {
	id: string;
	type: string;
	target: string;
	external: boolean;
	/** The element text, for rewriting. */
	text: string;
}

export const dirName = (path: string): string => {
	const i = path.lastIndexOf('/');
	return i < 0 ? '' : path.slice(0, i);
};
export const baseName = (path: string): string => path.slice(path.lastIndexOf('/') + 1);
export const relsPathOf = (part: string): string => {
	const dir = dirName(part);
	return `${dir ? `${dir}/` : ''}_rels/${baseName(part)}.rels`;
};

/** Resolves a relationship target against the directory of its source part. */
export function resolveTarget(dir: string, target: string): string {
	const parts = target.startsWith('/') ? [] : dir ? dir.split('/') : [];
	for (const segment of target.split('/')) {
		if (segment === '' || segment === '.') continue;
		if (segment === '..') parts.pop();
		else parts.push(segment);
	}
	return parts.join('/');
}

export function parseRelationships(xml: string | undefined): Relationship[] {
	if (!xml) return [];
	const out: Relationship[] = [];
	for (const match of xml.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)) {
		const text = match[0];
		const id = attrOf(text, 'Id');
		const type = attrOf(text, 'Type');
		const target = attrOf(text, 'Target');
		if (id === undefined || type === undefined || target === undefined) continue;
		out.push({ id, type, target, external: attrOf(text, 'TargetMode') === 'External', text });
	}
	return out;
}

export const PAGE_TYPE = 'http://schemas.microsoft.com/visio/2010/relationships/page';
export const PAGE_CONTENT_TYPE = 'application/vnd.ms-visio.page+xml';
export const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Parts that belong to the person's own window, never shared. */
export const isLocalOnlyPart = (path: string): boolean =>
	/(^|\/)windows\.xml$/i.test(path) || /^docProps\/thumbnail\./i.test(path);

export const CONTENT_TYPES = '[Content_Types].xml';

/** The content types part without the page overrides, which depend on each peer's part names. */
export function stripPageOverrides(xml: string): string {
	return xml.replace(/<(?:\w+:)?Override\b[^>]*>/g, (tag) =>
		attrOf(tag, 'ContentType') === PAGE_CONTENT_TYPE ? '' : tag,
	);
}

/** Adds an override for each page part, replacing any it already declares. */
export const withPageOverrides = (xml: string, pagePaths: readonly string[]): string =>
	withOverrides(stripPageOverrides(xml), PAGE_CONTENT_TYPE, pagePaths);

export const MASTER_TYPE = 'http://schemas.microsoft.com/visio/2010/relationships/master';
export const MASTER_CONTENT_TYPE = 'application/vnd.ms-visio.master+xml';

/** Appends an override of `contentType` for each of `paths`. */
export function withOverrides(xml: string, contentType: string, paths: readonly string[]): string {
	const add = paths
		.map((path) => `<Override PartName="/${path}" ContentType="${contentType}"/>`)
		.join('');
	const close = xml.lastIndexOf('</');
	return close < 0 ? xml : xml.slice(0, close) + add + xml.slice(close);
}

/** Removes the overrides of `types` and of the given part names. */
export function stripOverrides(
	xml: string,
	types: readonly string[],
	parts: ReadonlySet<string> = new Set(),
): string {
	return xml.replace(/<(?:\w+:)?Override\b[^>]*>/g, (tag) => {
		const type = attrOf(tag, 'ContentType');
		const name = attrOf(tag, 'PartName');
		return (type !== undefined && types.includes(type)) ||
			(name !== undefined && parts.has(name.replace(/^\//, '')))
			? ''
			: tag;
	});
}

export interface DeclaredType {
	type: string;
	/** True when it comes from a `Default` extension entry rather than an `Override`. */
	byDefault: boolean;
}
/** The content type the types part declares for `path`. */
export function declaredType(xml: string, path: string): DeclaredType | undefined {
	for (const tag of xml.matchAll(/<(?:\w+:)?Override\b[^>]*>/g)) {
		const name = attrOf(tag[0], 'PartName');
		const type = attrOf(tag[0], 'ContentType');
		if (name === `/${path}` && type !== undefined) return { type, byDefault: false };
	}
	const ext = /\.([^./]+)$/.exec(path)?.[1]?.toLowerCase();
	if (ext === undefined) return undefined;
	for (const tag of xml.matchAll(/<(?:\w+:)?Default\b[^>]*>/g)) {
		const e = attrOf(tag[0], 'Extension');
		const type = attrOf(tag[0], 'ContentType');
		if (e?.toLowerCase() === ext && type !== undefined) return { type, byDefault: true };
	}
	return undefined;
}

/** Declares `declared` for `path` unless the types part already covers it. */
export function ensureType(xml: string, path: string, declared: DeclaredType): string {
	if (declaredType(xml, path)) return xml;
	const close = xml.lastIndexOf('</');
	if (close < 0) return xml;
	const ext = /\.([^./]+)$/.exec(path)?.[1];
	const tag =
		declared.byDefault && ext
			? `<Default Extension="${ext}" ContentType="${declared.type}"/>`
			: `<Override PartName="/${path}" ContentType="${declared.type}"/>`;
	return xml.slice(0, close) + tag + xml.slice(close);
}

/** True for a package part directly inside a `media` folder. */
export const isMediaPart = (path: string): boolean => /(^|\/)media\/[^/]+$/i.test(path);

/** Relative target from the directory `dir` to the part `path`. */
export function relativeTarget(dir: string, path: string): string {
	const from = dir ? dir.split('/') : [];
	const to = path.split('/');
	let i = 0;
	while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
	return [...from.slice(i).map(() => '..'), ...to.slice(i)].join('/');
}

export function toBase64(bytes: Uint8Array): string {
	let out = '';
	for (let i = 0; i < bytes.length; i += 0x8000)
		out += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
	return btoa(out);
}
export function fromBase64(text: string): Uint8Array {
	const raw = atob(text);
	const out = new Uint8Array(raw.length);
	for (let i = 0; i < raw.length; i++) out[i] = raw.charCodeAt(i);
	return out;
}

/** Length plus an FNV-1a hash: enough to skip re-encoding an unchanged part. */
export function fingerprint(bytes: Uint8Array): string {
	let hash = 0x811c9dc5;
	for (let i = 0; i < bytes.length; i++) hash = Math.imul(hash ^ bytes[i]!, 0x01000193);
	return `${bytes.length}:${(hash >>> 0).toString(16)}`;
}

export const encoder = new TextEncoder();
export const decoder = new TextDecoder();

const canonCache = new Map<string, string>();
/**
 * A form of an XML fragment that ignores how a writer happened to serialize it (quote style,
 * empty elements), so that comparing a shape with what it was before an editor rewrote its page
 * does not mistake the rewrite for an edit. Falls back to the text when it does not parse alone.
 */
export function canon(xml: string): string {
	const cached = canonCache.get(xml);
	if (cached !== undefined) return cached;
	let result = xml;
	try {
		result = buildXml(parseXml(xml).documentElement);
	} catch {
		/* A fragment with a prefix declared elsewhere compares as text. */
	}
	if (canonCache.size > 4000) canonCache.clear();
	canonCache.set(xml, result);
	return result;
}

/** SHA-256 of `bytes` as lower-case hex: the identity of a media part. */
export function contentHash(bytes: Uint8Array): string {
	let out = '';
	for (const byte of sha256(bytes)) out += byte.toString(16).padStart(2, '0');
	return out;
}
