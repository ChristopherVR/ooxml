// How ink looks on the editing surface outside Print Layout: its strokes as SVG at the drawing's
// size, placed where its wrapping puts it (a floating ink at its offsets, like a drawn shape).
import type { DocxInk, PicturePlacement } from 'ooxml-core/docx';
import { inkPaint } from 'ooxml-core/docx/layout';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { NodeView } from 'prosemirror-view';

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

/** Paints ink into `dom`: its SVG, size and floating position. */
export function paintInk(dom: HTMLElement, node: ProseMirrorNode, selected = false): void {
	const ink = inkOf(node);
	const width = Number(node.attrs.widthPx) || 1;
	const height = Number(node.attrs.heightPx) || 1;
	const placement = parse<PicturePlacement>(node.attrs.placement);
	// Word's ink floats in front of the text at its offsets; ink in line with text sits in the line.
	const floating = Boolean(node.attrs.anchored);
	dom.className = `dve-ink${floating ? ' dve-ink-floating' : ''}${selected ? ' dve-ink-selected' : ''}`;
	dom.style.left = floating ? `${placement?.offsetXPx ?? 0}px` : '';
	dom.style.top = floating ? `${placement?.offsetYPx ?? 0}px` : '';
	dom.style.width = `${width}px`;
	dom.style.height = `${height}px`;
	dom.setAttribute('role', 'img');
	dom.setAttribute('aria-label', String(node.attrs.altText || node.attrs.drawingName || 'Ink'));
	const markup = inkPaint({ ...(ink ? { ink } : {}), widthPx: width, heightPx: height }).chart?.svg;
	const svg = markup
		? new DOMParser().parseFromString(markup, 'image/svg+xml').documentElement
		: undefined;
	dom.replaceChildren(...(svg?.localName === 'svg' ? [document.importNode(svg, true)] : []));
}

export function inkNodeView(initial: ProseMirrorNode): NodeView {
	let node = initial;
	let selected = false;
	const dom = document.createElement('span');
	dom.contentEditable = 'false';
	dom.dataset.docxInk = '';
	paintInk(dom, node);
	return {
		dom,
		update(next) {
			if (next.type !== node.type || !next.attrs.ink) return false;
			node = next;
			paintInk(dom, node, selected);
			return true;
		},
		selectNode: () => {
			selected = true;
			dom.classList.add('dve-ink-selected');
		},
		deselectNode: () => {
			selected = false;
			dom.classList.remove('dve-ink-selected');
		},
		ignoreMutation: () => true,
	};
}
