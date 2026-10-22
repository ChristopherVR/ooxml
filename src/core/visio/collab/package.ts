// Package-level helpers of the Visio mapping: relationships, part paths, the parts that stay local,
// base64 and a cheap fingerprint.
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
export function withPageOverrides(xml: string, pagePaths: readonly string[]): string {
	const clean = stripPageOverrides(xml);
	const add = pagePaths
		.map((path) => `<Override PartName="/${path}" ContentType="${PAGE_CONTENT_TYPE}"/>`)
		.join('');
	const close = clean.lastIndexOf('</');
	return close < 0 ? clean : clean.slice(0, close) + add + clean.slice(close);
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
