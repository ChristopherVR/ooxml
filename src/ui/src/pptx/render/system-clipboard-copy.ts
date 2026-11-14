/**
 * Copy shapes to the system clipboard as well as the viewer's own clipboard: PowerPoint offers
 * other applications a picture of the selection and its text, so a copied shape can be pasted
 * into a chat, a mail or an image editor. The viewer keeps its in-memory element snapshot for
 * pasting back into a deck; this module only adds the `image/png` and `text/plain` formats.
 *
 * Best effort and quiet: a browser without the asynchronous Clipboard API, a denied permission
 * or a failed rendering never throws and never blocks the internal copy. Nothing is sent
 * anywhere; the image is rendered locally from the mounted slide.
 *
 * @module render/system-clipboard-copy
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { findCanvasElementNode } from './canvas-element-node';
import { collectElementText } from './slide-search';

/** What reached the system clipboard. */
export type SystemClipboardCopyStatus = 'image-and-text' | 'text' | 'unavailable' | 'failed';

/** The clipboard surface used (a test seam; defaults to `navigator.clipboard`). */
export interface SystemClipboardTarget {
	write?: (items: ClipboardItem[]) => Promise<void>;
	writeText?: (text: string) => Promise<void>;
}

interface ClipboardItemConstructor {
	new (items: Record<string, Blob | PromiseLike<Blob>>): ClipboardItem;
	supports?: (type: string) => boolean;
}

/** The plain text PowerPoint puts beside copied shapes: each shape's text, one per line. */
export function elementsClipboardText(elements: readonly PptxElement[]): string {
	return elements.map(collectElementText).filter(Boolean).join('\n');
}

/**
 * Write `text` and, when given, the PNG `image` (a promise, so the write starts inside the user
 * gesture while the picture renders) to the system clipboard. Falls back to text alone when the
 * image is unsupported or fails.
 */
export async function writeSystemClipboard(
	input: { text: string; image?: Promise<Blob | null> },
	target: SystemClipboardTarget | undefined = globalThis.navigator?.clipboard,
	Item: ClipboardItemConstructor | undefined = (
		globalThis as { ClipboardItem?: ClipboardItemConstructor }
	).ClipboardItem,
): Promise<SystemClipboardCopyStatus> {
	// A rejected image must not surface as an unhandled rejection when it is never awaited.
	input.image?.catch(() => null);
	if (!target) {
		return 'unavailable';
	}
	const writeText = async (): Promise<SystemClipboardCopyStatus> => {
		if (!input.text) {
			return 'failed';
		}
		try {
			if (Item && target.write) {
				await target.write([
					new Item({ 'text/plain': new Blob([input.text], { type: 'text/plain' }) }),
				]);
			} else if (target.writeText) {
				await target.writeText(input.text);
			} else {
				return 'unavailable';
			}
			return 'text';
		} catch {
			return 'failed';
		}
	};
	const imageSupported = Item && target.write && (Item.supports?.('image/png') ?? true);
	if (!input.image || !imageSupported) {
		return writeText();
	}
	const image = input.image.then((blob) => {
		if (!blob) {
			throw new Error('The selection could not be rendered');
		}
		return blob.type === 'image/png' ? blob : new Blob([blob], { type: 'image/png' });
	});
	const formats: Record<string, Blob | PromiseLike<Blob>> = { 'image/png': image };
	if (input.text) {
		formats['text/plain'] = new Blob([input.text], { type: 'text/plain' });
	}
	try {
		await target.write!([new Item!(formats)]);
		return 'image-and-text';
	} catch {
		return writeText();
	}
}

/** Renders one mounted element node to a PNG data URL (each binding's picture rasteriser). */
export type ElementNodeRasterizer = (
	node: HTMLElement,
	elementId: string,
) => Promise<string | null>;

async function dataUrlToBlob(dataUrl: string): Promise<Blob | null> {
	const match = /^data:image\/png;base64,(.*)$/s.exec(dataUrl);
	if (!match) {
		return null;
	}
	const binary = atob(match[1] ?? '');
	const bytes = new Uint8Array(binary.length);
	for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
	return new Blob([bytes], { type: 'image/png' });
}

function loadImage(doc: Document, src: string): Promise<HTMLImageElement> {
	return new Promise((resolve, reject) => {
		const image = doc.createElement('img');
		image.onload = () => resolve(image);
		image.onerror = () => reject(new Error('image'));
		image.src = src;
	});
}

/**
 * A PNG of the selected elements as they are drawn on the canvas: one rendering per element,
 * placed by its on-screen box into the selection's bounding box. `null` when nothing is mounted.
 */
export async function renderSelectionPng(
	doc: Document,
	elementIds: readonly string[],
	rasterize: ElementNodeRasterizer,
	scale = 2,
): Promise<Blob | null> {
	const found = elementIds.flatMap((id) => {
		const node = findCanvasElementNode(doc, id, { canvasOnly: true });
		return node ? [{ id, node }] : [];
	});
	const nodes = found.map((entry) => entry.node);
	const ids = found.map((entry) => entry.id);
	if (!nodes.length) {
		return null;
	}
	if (nodes.length === 1) {
		const url = await rasterize(nodes[0] as HTMLElement, ids[0] as string);
		return url ? dataUrlToBlob(url) : null;
	}
	const rects = nodes.map((node) => node.getBoundingClientRect());
	const left = Math.min(...rects.map((r) => r.left));
	const top = Math.min(...rects.map((r) => r.top));
	const width = Math.max(...rects.map((r) => r.right)) - left;
	const height = Math.max(...rects.map((r) => r.bottom)) - top;
	const canvas = doc.createElement('canvas');
	canvas.width = Math.max(1, Math.round(width * scale));
	canvas.height = Math.max(1, Math.round(height * scale));
	const context = canvas.getContext('2d');
	if (!context) {
		return null;
	}
	for (const [index, node] of nodes.entries()) {
		const url = await rasterize(node, ids[index] as string);
		const rect = rects[index] as DOMRect;
		if (!url) continue;
		const image = await loadImage(doc, url);
		context.drawImage(
			image,
			(rect.left - left) * scale,
			(rect.top - top) * scale,
			rect.width * scale,
			rect.height * scale,
		);
	}
	return new Promise((resolve) => canvas.toBlob((blob) => resolve(blob), 'image/png'));
}

/**
 * The system half of Copy for slide elements: their text and, when a rasteriser is given, a
 * picture of them. Call it from the copy command, after the internal snapshot is taken.
 */
export function copyElementsToSystemClipboard(input: {
	elements: readonly PptxElement[];
	doc?: Document;
	rasterize?: ElementNodeRasterizer;
	target?: SystemClipboardTarget;
}): Promise<SystemClipboardCopyStatus> {
	const doc = input.doc ?? globalThis.document;
	const text = elementsClipboardText(input.elements);
	const image =
		doc && input.rasterize && input.elements.length
			? renderSelectionPng(
					doc,
					input.elements.map((element) => element.id),
					input.rasterize,
				).catch(() => null)
			: undefined;
	return writeSystemClipboard(
		{ text, ...(image ? { image } : {}) },
		'target' in input ? input.target : globalThis.navigator?.clipboard,
	);
}
