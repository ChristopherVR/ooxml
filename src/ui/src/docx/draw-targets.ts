// The ink drawings the Draw tab works on, found on the Print Layout pages: floating ink (a page
// float) and ink in line with text (a picture on a line), each with its node, its page and the
// box it is drawn in. Ink is told apart by the model (`data-ink`), never by its label.
import type { InkCoordinate } from 'ooxml-core/drawingml';
import type { DocxInk } from 'ooxml-core/docx';
import type { EditorView } from 'prosemirror-view';
import { paragraphStart } from './ink-commands';
import { drawHelpers } from './draw-helpers';

/** One ink drawing on a page, with the strokes the lasso holds in it. */
export interface InkTarget {
	pos: number;
	ink: DocxInk;
	/** The node's `ink` attribute when it was found: a later edit of the drawing invalidates it. */
	attr: string;
	element: HTMLElement;
	sheet: HTMLElement;
	blockId: string;
	indexes: number[];
}

/** The zoom of a page: its drawn width over its layout width. */
export const scaleOf = (element: HTMLElement): number =>
	(element.offsetWidth ? element.getBoundingClientRect().width / element.offsetWidth : 1) || 1;

/**
 * The ink drawings on the pages whose box `test` accepts. `editable` leaves out ink read only in
 * part (`DocxInk.truncated`), whose strokes the tools must not change.
 */
export function inkTargets(
	view: EditorView,
	pages: HTMLElement | undefined,
	test: (rect: DOMRect) => boolean,
	editable: boolean,
): InkTarget[] {
	if (!pages) return [];
	return [...pages.querySelectorAll<HTMLElement>('.dve-print-page [data-ink]')].flatMap(
		(element) => {
			const sheet = element.closest<HTMLElement>('.dve-print-page');
			const line = element.closest<HTMLElement>('.dve-print-line[data-block-id]');
			const blockId = element.dataset.blockId ?? line?.dataset.blockId;
			if (!sheet || !blockId || !test(element.getBoundingClientRect())) return [];
			const pos = element.classList.contains('dve-print-float')
				? drawHelpers().locateFloat(view.state.doc, blockId, Number(element.dataset.floatIndex))
						?.pos
				: inlineAt(view, blockId, Number(element.dataset.sourceStart));
			const node = pos === undefined ? null : view.state.doc.nodeAt(pos);
			const ink = node ? drawHelpers().inkOf(node) : undefined;
			if (pos === undefined || !node || !ink || (editable && ink.truncated)) return [];
			return [{ pos, ink, attr: String(node.attrs.ink), element, sheet, blockId, indexes: [] }];
		},
	);
}

/** The position of a picture on a line: its paragraph's start and its source offset there. */
function inlineAt(view: EditorView, blockId: string, offset: number): number | undefined {
	const start = paragraphStart(view.state.doc, blockId);
	return start === undefined ? undefined : start + offset;
}

/** Whether the drawing a target was found on is still there, unchanged. */
export const stillThere = (view: EditorView, target: InkTarget): boolean =>
	view.state.doc.nodeAt(target.pos)?.attrs.ink === target.attr;

/** A client point in an ink drawing's own coordinates (and the factor for lengths). */
export function localPoint(
	target: InkTarget,
	point: InkCoordinate,
): { point: InkCoordinate; factor: number } {
	const rect = target.element.getBoundingClientRect();
	const factor = target.ink.width / (rect.width || 1);
	return {
		point: {
			x: (point.x - rect.left) * factor,
			y: (point.y - rect.top) * (target.ink.height / (rect.height || 1)),
		},
		factor,
	};
}

/** How a target's own coordinates map onto its page (CSS pixels from the page's corner). */
export function pageMapping(target: InkTarget): { ox: number; oy: number; kx: number; ky: number } {
	const rect = target.element.getBoundingClientRect();
	const sheet = target.sheet.getBoundingClientRect();
	const scale = scaleOf(target.sheet);
	return {
		ox: (rect.left - sheet.left) / scale,
		oy: (rect.top - sheet.top) / scale,
		kx: rect.width / (target.ink.width || 1) / scale,
		ky: rect.height / (target.ink.height || 1) / scale,
	};
}

/** The page under a client point. */
export function sheetAt(
	pages: HTMLElement | undefined,
	point: InkCoordinate,
): HTMLElement | undefined {
	return [...(pages?.querySelectorAll<HTMLElement>('.dve-print-page') ?? [])].find((sheet) => {
		const rect = sheet.getBoundingClientRect();
		return (
			point.x >= rect.left && point.x <= rect.right && point.y >= rect.top && point.y <= rect.bottom
		);
	});
}

/**
 * The body paragraph a drawing at `top` (client) on `sheet` anchors to: the last line above it, or
 * the page's first. Only paragraphs of the body itself: a float anchored in a table cell is laid
 * out from the cell, so page offsets would put it somewhere else.
 */
export function bodyAnchor(view: EditorView, sheet: HTMLElement, top: number): string | undefined {
	let chosen: string | undefined;
	let first: string | undefined;
	for (const line of sheet.querySelectorAll<HTMLElement>('.dve-print-line[data-block-id]')) {
		const id = line.dataset.blockId!;
		const at = paragraphStart(view.state.doc, id);
		if (at === undefined || view.state.doc.resolve(at).depth !== 1) continue;
		first ??= id;
		if (line.getBoundingClientRect().top <= top) chosen = id;
	}
	return chosen ?? first;
}
