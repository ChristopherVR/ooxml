// How ink looks on the editing surface outside Print Layout: its strokes as SVG at the drawing's
// size, placed where its wrapping puts it (a floating ink at its offsets, like a drawn shape).
// Loaded with the first ink a document shows (`ink-view.ts`).
import type { DocxInk, LegacyVmlInk, PicturePlacement } from 'ooxml-core/docx';
import { inkPaint } from 'ooxml-core/docx/layout';
import type { Node as ProseMirrorNode } from 'prosemirror-model';

function parse<T>(value: unknown): T | undefined {
	if (typeof value !== 'string') return undefined;
	try {
		return JSON.parse(value) as T;
	} catch {
		return undefined;
	}
}

/** The ink of an image node, when it is ink. */
export const inkOf = (node: ProseMirrorNode): DocxInk | undefined =>
	node.type.name === 'image' ? parse<DocxInk>(node.attrs.ink) : undefined;

/** Paints ink into `dom`: its SVG, size and floating position (a selected drawing stays selected). */
export function paintInk(dom: HTMLElement, node: ProseMirrorNode): void {
	const width = Number(node.attrs.widthPx) || 1;
	const height = Number(node.attrs.heightPx) || 1;
	// Word's ink floats in front of the text at its offsets; ink in line with text sits in the line.
	const at = node.attrs.anchored
		? (parse<PicturePlacement>(node.attrs.placement) ?? ({} as PicturePlacement))
		: undefined;
	const selected = dom.classList.contains('dve-ink-selected');
	dom.className = at ? 'dve-ink dve-ink-floating' : 'dve-ink';
	dom.classList.toggle('dve-ink-selected', selected);
	Object.assign(dom.style, {
		left: at ? `${at.offsetXPx ?? 0}px` : '',
		top: at ? `${at.offsetYPx ?? 0}px` : '',
		width: `${width}px`,
		height: `${height}px`,
	});
	dom.setAttribute('role', 'img');
	// Ink drawn here is named in the editor's language when it is drawn (`insertInk`); Word names
	// its own (`Ink 1`).
	dom.setAttribute('aria-label', String(node.attrs.altText || node.attrs.drawingName || 'Ink'));
	// An unchanged drawing at an unchanged size is not drawn again.
	const host = dom as HTMLElement & { _ink?: unknown };
	const key = `${width} ${height} ${node.attrs.ink ?? node.attrs.legacyInk}`;
	if (host._ink === key) return;
	host._ink = key;
	const ink = inkOf(node);
	// Legacy VML ink is drawn the same way, read only: the Draw tab finds only `ink`.
	const legacyInk = ink ? undefined : parse<LegacyVmlInk>(node.attrs.legacyInk);
	if (legacyInk) dom.dataset.legacyInk = '';
	else delete dom.dataset.legacyInk;
	const markup = inkPaint({
		...(ink ? { ink } : {}),
		...(legacyInk ? { legacyInk } : {}),
		widthPx: width,
		heightPx: height,
	}).chart?.svg;
	const svg = markup
		? new DOMParser().parseFromString(markup, 'image/svg+xml').documentElement
		: undefined;
	dom.replaceChildren(...(svg?.localName === 'svg' ? [document.importNode(svg, true)] : []));
}
