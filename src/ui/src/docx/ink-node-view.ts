// How ink looks on the editing surface outside Print Layout: its strokes as SVG at the drawing's
// size, placed where its wrapping puts it (a floating ink at its offsets, like a drawn shape).
import type { DocxInk, PicturePlacement } from 'ooxml-core/docx';
import { inkPaint } from 'ooxml-core/docx/layout';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { NodeView } from 'prosemirror-view';
import { translate, type EditorLocale } from './localization';

function parse<T>(value: unknown): T | undefined {
	if (typeof value !== 'string') return undefined;
	try {
		return JSON.parse(value) as T;
	} catch {
		return undefined;
	}
}

/** Parsed `ink` attributes by text: every tool and paint reads the same drawing again and again. */
const parsed = new Map<string, DocxInk | undefined>();

/** The ink of an image node, when it is ink. Shared between callers: do not change it. */
export function inkOf(node: ProseMirrorNode): DocxInk | undefined {
	const attr = node.type.name === 'image' ? node.attrs.ink : undefined;
	if (typeof attr !== 'string') return undefined;
	if (parsed.has(attr)) return parsed.get(attr);
	if (parsed.size >= 64) parsed.delete(parsed.keys().next().value as string);
	const ink = parse<DocxInk>(attr);
	parsed.set(attr, ink);
	return ink;
}

/** What `dom` was last painted from, so an unchanged drawing is not painted again. */
const painted = new WeakMap<HTMLElement, string>();

/** Paints ink into `dom`: its SVG, size and floating position. */
export function paintInk(
	dom: HTMLElement,
	node: ProseMirrorNode,
	selected = false,
	locale: EditorLocale = 'en',
): void {
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
	const name = node.attrs.drawingName === 'Ink' ? '' : node.attrs.drawingName;
	dom.setAttribute(
		'aria-label',
		String(node.attrs.altText || name || translate(locale, 'draw.ink')),
	);
	const key = `${width}x${height}:${String(node.attrs.ink)}`;
	if (painted.get(dom) === key) return;
	painted.set(dom, key);
	const markup = inkPaint({ ...(ink ? { ink } : {}), widthPx: width, heightPx: height }).chart?.svg;
	const svg = markup
		? new DOMParser().parseFromString(markup, 'image/svg+xml').documentElement
		: undefined;
	dom.replaceChildren(...(svg?.localName === 'svg' ? [document.importNode(svg, true)] : []));
}

export function inkNodeView(initial: ProseMirrorNode, locale?: EditorLocale): NodeView {
	let node = initial;
	let selected = false;
	const dom = document.createElement('span');
	dom.contentEditable = 'false';
	dom.dataset.docxInk = '';
	paintInk(dom, node, false, locale);
	return {
		dom,
		update(next) {
			if (next.type !== node.type || !next.attrs.ink) return false;
			node = next;
			paintInk(dom, node, selected, locale);
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
