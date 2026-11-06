/**
 * Per-copy ids for SVG defs.
 *
 * A slide is painted several times on one page (the editor canvas, thumbnails,
 * the slide show, morph overlay copies, reflection mirrors) and every renderer
 * derives the ids of its filters, markers, gradients, patterns, clip paths and
 * masks from the element id alone. `url(#id)` resolves to the FIRST element with
 * the id in the document, so a copy painted inside a hidden or differently
 * styled subtree supplied the def for all the others: arrowheads vanished in a
 * show, silhouette shadows painted at the default flood opacity.
 *
 * One scheme for every binding: `installSvgIdScope(document)` watches the
 * document, gives every referenced def its own id (`<id>~<n>`, the original kept
 * in `data-pptx-sid`) and points each reference at the def its own copy
 * contains: the nearest one found walking up from the reference. Renderers and
 * bindings keep deriving ids from element ids and need no per-copy plumbing.
 */
import { escapeCssAttributeValue } from './css-escape';

const SID = 'data-pptx-sid';
const SCOPED_SUFFIX = /~\d+$/u;
const URL_REF = /url\(\s*(['"]?)#([^'")\s]+)\1\s*\)/gu;
/** The attributes that reference a def with `url(#id)`. */
const URL_ATTRS = [
	'filter',
	'fill',
	'stroke',
	'clip-path',
	'mask',
	'marker-start',
	'marker-mid',
	'marker-end',
	'style',
];
/** SVG elements that reference a def with `href="#id"`. */
const HREF_TAGS = ['use', 'textPath', 'linearGradient', 'radialGradient', 'pattern', 'feImage'];
const HREF_ATTRS = ['href', 'xlink:href'];

const REF_SELECTOR = [
	...URL_ATTRS.map((a) => `[${a}*="url("]`),
	...HREF_TAGS.flatMap((tag) => HREF_ATTRS.map((a) => `${tag}[${a.replace(':', '\\:')}^="#"]`)),
].join(',');

const OBSERVED_ATTRS = ['id', ...URL_ATTRS, ...HREF_ATTRS];

/** Counts scoped defs; one counter per page keeps every id unique. */
let serial = 0;
/** The id each def was last given, so a def a framework re-ids is noticed. */
const owned = new WeakMap<Element, string>();

function originalId(token: string): string {
	return token.replace(SCOPED_SUFFIX, '');
}

function findDef(from: Element, original: string): Element | undefined {
	const selector = `[id="${escapeCssAttributeValue(original)}"],[${SID}="${escapeCssAttributeValue(original)}"]`;
	for (let scope: Element | null = from; scope; scope = scope.parentElement) {
		const hit = scope.querySelector(selector);
		if (hit) {
			return hit;
		}
	}
	return undefined;
}

/** Gives `def` an id of its own (once) and returns it. */
function scopeDef(def: Element): string {
	const current = def.id;
	if (owned.get(def) !== current) {
		const original = originalId(current);
		def.setAttribute(SID, original);
		def.id = `${original}~${++serial}`;
		owned.set(def, def.id);
	}
	return def.id;
}

function retarget(node: Element, token: string): string {
	const def = findDef(node, originalId(token));
	return def ? scopeDef(def) : token;
}

function scopeUrls(node: Element, attr: string): void {
	const value = node.getAttribute(attr);
	if (!value || !value.includes('url(')) {
		return;
	}
	const next = value.replace(URL_REF, (whole, quote: string, token: string) => {
		const id = retarget(node, token);
		return id === token ? whole : `url(${quote}#${id}${quote})`;
	});
	if (next !== value) {
		node.setAttribute(attr, next);
	}
}

function scopeHref(node: Element, attr: string): void {
	const value = node.getAttribute(attr);
	if (!value || !value.startsWith('#')) {
		return;
	}
	const id = retarget(node, value.slice(1));
	if (`#${id}` !== value) {
		node.setAttribute(attr, `#${id}`);
	}
}

function scopeReference(node: Element): void {
	for (const attr of URL_ATTRS) {
		scopeUrls(node, attr);
	}
	if (HREF_TAGS.includes(node.localName)) {
		for (const attr of HREF_ATTRS) {
			scopeHref(node, attr);
		}
	}
}

/** Scopes every reference inside `root` (and `root` itself). */
export function scopeSvgIds(root: Element | Document): void {
	if (root.nodeType === 1 && (root as Element).matches(REF_SELECTOR)) {
		scopeReference(root as Element);
	}
	for (const node of root.querySelectorAll(REF_SELECTOR)) {
		scopeReference(node);
	}
}

interface Installed {
	count: number;
	observer: MutationObserver;
}

const installed = new WeakMap<Document, Installed>();

function dirtyRoot(node: Node): Element | undefined {
	const element = node.nodeType === 1 ? (node as Element) : node.parentElement;
	return element?.closest('[data-element-id]') ?? element?.parentElement ?? element ?? undefined;
}

function relevant(record: MutationRecord): boolean {
	if (record.type === 'childList') {
		return record.addedNodes.length > 0;
	}
	const target = record.target as Element;
	if (record.attributeName === 'id') {
		return true;
	}
	return (target.getAttribute(record.attributeName ?? '') ?? '').includes('#');
}

/**
 * Starts scoping the ids of `doc` (idempotent, reference counted). Returns a
 * function that releases it.
 */
export function installSvgIdScope(doc: Document = document): () => void {
	const existing = installed.get(doc);
	if (existing) {
		existing.count += 1;
		return () => release(doc);
	}
	const view = doc.defaultView;
	if (!view || !doc.documentElement) {
		return () => undefined;
	}
	const observer = new view.MutationObserver((records) => {
		const roots = new Set<Element>();
		for (const record of records) {
			if (!relevant(record)) {
				continue;
			}
			if (record.type === 'childList') {
				for (const added of record.addedNodes) {
					const root = dirtyRoot(added);
					if (root) {
						roots.add(root);
					}
				}
			} else {
				const root = dirtyRoot(record.target);
				if (root) {
					roots.add(root);
				}
			}
		}
		for (const root of roots) {
			if (root.isConnected) {
				scopeSvgIds(root);
			}
		}
		observer.takeRecords();
	});
	observer.observe(doc.documentElement, {
		subtree: true,
		childList: true,
		attributes: true,
		attributeFilter: OBSERVED_ATTRS,
	});
	installed.set(doc, { count: 1, observer });
	scopeSvgIds(doc.documentElement);
	return () => release(doc);
}

function release(doc: Document): void {
	const entry = installed.get(doc);
	if (!entry) {
		return;
	}
	entry.count -= 1;
	if (entry.count <= 0) {
		entry.observer.disconnect();
		installed.delete(doc);
	}
}
