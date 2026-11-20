// Chromium's style cache keeps the computed style of a child whose parent references a
// `url(#id)` clip path, mask, filter, marker or paint server (`fill="url(#g)"` on a group, say).
// That style points at the tree the reference resolves in, so a removed editor whose shadow root
// painted such a group stays alive, with its document, until the document's styles next change.
// Every stylesheet change empties that cache, so a removed editor makes one that matches nothing.

const pending = new WeakMap<Document, Set<Element>>();
const sheets = new WeakMap<Document, CSSStyleSheet>();

/**
 * Lets the browser drop the cached styles of `element`'s tree once the current task ends, when
 * `element` is still detached by then (a move between parents keeps its styles). Calls made in
 * one task share a single stylesheet change. Call it from `disconnectedCallback`.
 */
export function releaseDetachedStyles(element: Element): void {
	const doc = element.ownerDocument;
	const view = doc.defaultView;
	if (!view) return;
	let removed = pending.get(doc);
	if (!removed) {
		const waiting = new Set<Element>();
		pending.set(doc, (removed = waiting));
		view.setTimeout(() => {
			pending.delete(doc);
			if ([...waiting].some((node) => !node.isConnected)) changeStyles(doc, view);
		});
	}
	removed.add(element);
}

/** Toggles a rule that matches nothing in a stylesheet of the document's own. */
function changeStyles(doc: Document, view: Window & typeof globalThis): void {
	if (!('adoptedStyleSheets' in doc) || typeof view.CSSStyleSheet !== 'function') return;
	let sheet = sheets.get(doc);
	if (!sheet) {
		try {
			sheet = new view.CSSStyleSheet();
		} catch {
			// No constructable stylesheets (jsdom): nothing caches styles there either.
			return;
		}
		sheets.set(doc, sheet);
	}
	// A host may have replaced the list since the last release.
	if (!doc.adoptedStyleSheets.includes(sheet))
		doc.adoptedStyleSheets = [...doc.adoptedStyleSheets, sheet];
	sheet.replaceSync(sheet.cssRules.length ? '' : ':not(*) {}');
}
