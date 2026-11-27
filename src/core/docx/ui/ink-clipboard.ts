// Ink through the clipboard: a copied or cut ink drawing carries its strokes, size, position and
// name in data attributes of its placeholder, so pasting it (here or in another editor) brings the
// ink back. The pasted ink has no part of its own yet (no relationship, no part name): the save
// writes it a new one, so a copy never shares the original's part.
import type { Node as ProseMirrorNode } from 'prosemirror-model';

/** Clipboard HTML of other origins may carry anything: past this size it is not read as ink. */
const MAX_CLIPBOARD_INK = 4 * 1024 * 1024;

/** The data attributes an ink drawing's placeholder carries. */
export function inkClipboardAttrs(node: ProseMirrorNode): Record<string, string> {
	if (typeof node.attrs.ink !== 'string') return {};
	let ink: Record<string, unknown>;
	try {
		ink = JSON.parse(node.attrs.ink) as Record<string, unknown>;
	} catch {
		return {};
	}
	const { relId: _relId, partName: _partName, truncated: _truncated, ...own } = ink;
	return {
		'data-ink': JSON.stringify(own),
		'data-unsupported': String(node.attrs.unsupported ?? 'Ink'),
		'data-width-px': String(node.attrs.widthPx),
		'data-height-px': String(node.attrs.heightPx),
		'data-anchored': node.attrs.anchored ? '1' : '0',
		...(typeof node.attrs.placement === 'string' ? { 'data-placement': node.attrs.placement } : {}),
		...(node.attrs.drawingName ? { 'data-drawing-name': String(node.attrs.drawingName) } : {}),
	};
}

const isStroke = (value: unknown): boolean =>
	typeof value === 'object' &&
	value !== null &&
	typeof (value as { path?: unknown }).path === 'string' &&
	typeof (value as { color?: unknown }).color === 'string' &&
	Number.isFinite((value as { width?: unknown }).width);

/** The ink node attributes read back from a pasted placeholder, or none when it holds no ink. */
export function inkFromClipboard(el: HTMLElement): Record<string, unknown> {
	const text = el.dataset.ink;
	if (!text || text.length > MAX_CLIPBOARD_INK) return {};
	let ink: { strokes?: unknown; width?: unknown; height?: unknown };
	try {
		ink = JSON.parse(text) as typeof ink;
	} catch {
		return {};
	}
	if (
		!Array.isArray(ink.strokes) ||
		!ink.strokes.every(isStroke) ||
		!((ink.width as number) > 0) ||
		!((ink.height as number) > 0)
	)
		return {};
	const placement = el.dataset.placement;
	let floating: unknown;
	try {
		floating = placement ? JSON.parse(placement) : undefined;
	} catch {
		floating = undefined;
	}
	return {
		ink: JSON.stringify({ strokes: ink.strokes, width: ink.width, height: ink.height }),
		...(floating && typeof floating === 'object' ? { placement } : {}),
		drawingName: el.dataset.drawingName || 'Ink',
	};
}
