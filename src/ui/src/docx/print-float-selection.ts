import { NodeSelection, Selection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
	CURSORS,
	boxOf,
	isFloat,
	isShapeFloat,
	locateFloat,
	number,
	parsePlacement,
	place,
	selectedFloat,
	shapeSpecOf,
	type Located,
} from './print-float-locate';
import { editShapeText } from './print-float-text';
import {
	FLOAT_HANDLES,
	NUDGE_PX,
	moveBox,
	placementAt,
	resizeBox,
	type FloatBox,
	type FloatHandle,
} from './print-float-geometry';

export { locateFloat, selectedFloat } from './print-float-locate';

export interface PrintFloatHost {
	/** The `.dve-print-pages` element. */
	pages: HTMLElement;
	/** The editing surface whose keys a selected float takes over. */
	paper: HTMLElement;
	view(): EditorView | undefined;
	printActive(): boolean;
}

export interface PrintFloats {
	/** Redraws the selection frame; call after transactions and layouts. */
	refresh(): void;
	destroy(): void;
}

/**
 * Floating pictures, shapes and text boxes on Print Layout pages are selectable like in Word: a press selects the object
 * (the same node selection the Draft surface uses) and draws a frame with eight handles; dragging
 * moves it, corner handles resize it (keeping the aspect ratio for pictures; a second press on a shape
 * that holds text, or a double-click on any shape, edits its text in place), arrow keys nudge it and Delete
 * removes it. Each gesture is one transaction, so it is one undo step. Pictures behind the text
 * are only taken where no text is, so a press on text still places the caret.
 */
export function attachPrintFloatSelection(host: PrintFloatHost): PrintFloats {
	const controller = new AbortController();
	let frame: HTMLElement | undefined;
	let frameFor: HTMLElement | undefined;
	let redraw: number | undefined;
	let closeText: (() => void) | undefined;

	/** Opens the in-place text editor on the selected shape or text box, when its text is editable. */
	function editText(located: Located, element: HTMLElement): void {
		const view = host.view();
		const spec = shapeSpecOf(located.node);
		if (!view || !spec?.textEditable || !spec.editable || !view.editable) return;
		closeText?.();
		closeText = editShapeText(view, element, spec.text ?? [], () => {
			closeText = undefined;
		});
	}

	const floatElement = (blockId: string, index: number): HTMLElement | undefined =>
		[...host.pages.querySelectorAll<HTMLElement>('.dve-print-float[data-block-id]')].find(
			(element) =>
				element.dataset.blockId === blockId && element.dataset.floatIndex === String(index),
		);

	const scaleOf = (sheet: HTMLElement): number =>
		(sheet.offsetWidth ? sheet.getBoundingClientRect().width / sheet.offsetWidth : 1) || 1;

	const originOf = (element: HTMLElement) => ({
		x: number(element.dataset.originX),
		y: number(element.dataset.originY),
	});

	function clearFrame(): void {
		frame?.remove();
		frame = undefined;
		frameFor = undefined;
	}

	function draw(): void {
		redraw = undefined;
		clearFrame();
		const view = host.view();
		if (!host.printActive() || !view || view.isDestroyed) return;
		const selected = selectedFloat(view.state);
		const element = selected && floatElement(selected.blockId, selected.index);
		const sheet = element?.closest<HTMLElement>('.dve-print-page');
		if (!element || !sheet) return;
		frame = document.createElement('div');
		frame.className = 'dve-print-float-frame';
		frame.setAttribute('aria-hidden', 'true');
		for (const handle of FLOAT_HANDLES) {
			const grip = document.createElement('span');
			grip.className = 'dve-print-float-handle';
			grip.dataset.floatHandle = handle;
			grip.style.cursor = CURSORS[handle];
			frame.append(grip);
		}
		place(frame, boxOf(element));
		frameFor = element;
		sheet.append(frame);
	}

	function schedule(): void {
		if (redraw === undefined) redraw = requestAnimationFrame(draw);
	}

	/** Draws the frame now, in place of a redraw waiting for the next animation frame. */
	function drawNow(): void {
		if (redraw !== undefined) cancelAnimationFrame(redraw);
		draw();
	}

	/** The float under a viewport point: in front of the text first, behind it only off the text. */
	function floatAt(x: number, y: number): HTMLElement | undefined {
		const inside = (element: HTMLElement) => {
			const rect = element.getBoundingClientRect();
			return x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom;
		};
		const floats = [
			...host.pages.querySelectorAll<HTMLElement>('.dve-print-float[data-block-id]'),
		].reverse();
		const front = floats.find(
			(element) => !element.classList.contains('dve-print-float-behind') && inside(element),
		);
		if (front) return front;
		const root = host.pages.getRootNode() as Document | ShadowRoot;
		const hit = root.elementFromPoint?.(x, y);
		if (hit?.closest('.dve-print-line')) return undefined;
		return floats.find((element) => inside(element));
	}

	function commit(located: Located, box: FloatBox, origin: { x: number; y: number }): void {
		const view = host.view();
		if (!view || view.isDestroyed || !view.editable) return;
		const current = view.state.doc.nodeAt(located.pos);
		if (!current || !isFloat(current)) return;
		// A shape whose outline or fill we do not model can be selected but not rewritten.
		if (isShapeFloat(current) && !shapeSpecOf(current)?.editable) return;
		const placement = placementAt(parsePlacement(current.attrs.placement), box, origin);
		const tr = view.state.tr.setNodeMarkup(located.pos, undefined, {
			...current.attrs,
			widthPx: Math.round(box.width),
			heightPx: Math.round(box.height),
			placement: JSON.stringify(placement),
		});
		tr.setSelection(NodeSelection.create(tr.doc, located.pos));
		view.dispatch(tr);
	}

	function select(element: HTMLElement): Located | undefined {
		const view = host.view();
		const blockId = element.dataset.blockId;
		if (!view || !blockId) return undefined;
		const located = locateFloat(view.state.doc, blockId, Number(element.dataset.floatIndex));
		if (!located) return undefined;
		const current = selectedFloat(view.state);
		if (current?.pos !== located.pos)
			view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, located.pos)));
		view.focus();
		return located;
	}

	function drag(
		event: PointerEvent,
		located: Located,
		element: HTMLElement,
		handle: FloatHandle | undefined,
		wasSelected: boolean,
	): void {
		const view = host.view();
		const sheet = element.closest<HTMLElement>('.dve-print-page');
		if (!view?.editable || !sheet) return;
		const reopen =
			!handle && wasSelected && Boolean(shapeSpecOf(located.node)?.text?.some(Boolean));
		const scale = scaleOf(sheet);
		const start = boxOf(element);
		const origin = originOf(element);
		const lockAspect = !isShapeFloat(located.node) && !located.node.attrs.unsupported;
		let box = start;
		let moved = false;
		const move = (moveEvent: PointerEvent) => {
			const dx = (moveEvent.clientX - event.clientX) / scale;
			const dy = (moveEvent.clientY - event.clientY) / scale;
			if (!moved && Math.hypot(dx, dy) < 3) return;
			moved = true;
			box = handle ? resizeBox(start, handle, dx, dy, lockAspect) : moveBox(start, dx, dy);
			place(element, box);
			if (frame) place(frame, box);
		};
		const finish = () => {
			window.removeEventListener('pointermove', move);
			window.removeEventListener('pointerup', finish);
			window.removeEventListener('pointercancel', finish);
			if (moved) commit(located, box, origin);
			else if (reopen) editText(located, element);
		};
		window.addEventListener('pointermove', move);
		window.addEventListener('pointerup', finish);
		window.addEventListener('pointercancel', finish);
	}

	host.pages.addEventListener(
		'pointerdown',
		(event) => {
			if (!host.printActive() || event.button !== 0 || !(event.target instanceof Element)) return;
			if (host.pages.classList.contains('dve-print-story-editing')) return;
			const grip = event.target.closest<HTMLElement>('[data-float-handle]');
			const element = grip ? frameFor : floatAt(event.clientX, event.clientY);
			if (!element) return;
			closeText?.();
			const view = host.view();
			const before = view && selectedFloat(view.state);
			const located = select(element);
			if (!located) return;
			event.preventDefault();
			event.stopImmediatePropagation();
			drag(
				event,
				located,
				element,
				grip?.dataset.floatHandle as FloatHandle | undefined,
				before?.pos === located.pos,
			);
		},
		{ capture: true, signal: controller.signal },
	);

	host.pages.addEventListener(
		'dblclick',
		(event) => {
			if (!host.printActive() || !host.view()) return;
			const element = floatAt(event.clientX, event.clientY);
			const located = element && select(element);
			if (!element || !located || !isShapeFloat(located.node)) return;
			event.preventDefault();
			event.stopImmediatePropagation();
			editText(located, element);
		},
		{ capture: true, signal: controller.signal },
	);

	host.paper.addEventListener(
		'keydown',
		(event) => {
			const view = host.view();
			if (!host.printActive() || !view || event.isComposing) return;
			const located = selectedFloat(view.state);
			if (!located) return;
			if (event.key === 'Escape') {
				const end = located.pos + located.node.nodeSize;
				view.dispatch(view.state.tr.setSelection(Selection.near(view.state.doc.resolve(end), 1)));
				event.preventDefault();
				event.stopImmediatePropagation();
				return;
			}
			const step = event.ctrlKey ? 1 : NUDGE_PX;
			const delta: Record<string, [number, number]> = {
				ArrowLeft: [-step, 0],
				ArrowRight: [step, 0],
				ArrowUp: [0, -step],
				ArrowDown: [0, step],
			};
			const move = delta[event.key];
			if (!move || event.altKey || event.metaKey || event.shiftKey) return;
			event.preventDefault();
			event.stopImmediatePropagation();
			const element = floatElement(located.blockId, located.index);
			if (!element || !view.editable) return;
			const box = moveBox(boxOf(element), move[0], move[1]);
			place(element, box);
			if (frame) place(frame, box);
			commit(located, box, originOf(element));
		},
		{ capture: true, signal: controller.signal },
	);

	// A relayout replaces the pages, the frame and its handles with them, and the page window mounts a
	// page again as it scrolls back into view. Draw the frame again before that task ends, not on the
	// next frame: a press in between would find no handle and, on a shape, start a move instead of a
	// resize. Two observers, because observing a node again replaces its earlier options.
	const relayout = new MutationObserver(drawNow);
	relayout.observe(host.pages, { childList: true });
	const remount = new MutationObserver(drawNow);
	remount.observe(host.pages, {
		attributes: true,
		attributeFilter: ['data-render-version'],
		subtree: true,
	});

	return {
		refresh: schedule,
		destroy() {
			controller.abort();
			closeText?.();
			relayout.disconnect();
			remount.disconnect();
			if (redraw !== undefined) cancelAnimationFrame(redraw);
			clearFrame();
		},
	};
}
