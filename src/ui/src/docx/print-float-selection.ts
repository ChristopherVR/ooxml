import { NodeSelection, Selection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
	CURSORS,
	boxOf,
	isFloat,
	locateFloat,
	number,
	parsePlacement,
	place,
	selectedFloat,
	type Located,
} from './print-float-locate';
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
 * Floating pictures on Print Layout pages are selectable like in Word: a press selects the picture
 * (the same node selection the Draft surface uses) and draws a frame with eight handles; dragging
 * moves it, corner handles resize it keeping the aspect ratio, arrow keys nudge it and Delete
 * removes it. Each gesture is one transaction, so it is one undo step. Pictures behind the text
 * are only taken where no text is, so a press on text still places the caret.
 */
export function attachPrintFloatSelection(host: PrintFloatHost): PrintFloats {
	const controller = new AbortController();
	let frame: HTMLElement | undefined;
	let frameFor: HTMLElement | undefined;
	let redraw: number | undefined;

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
	): void {
		const view = host.view();
		const sheet = element.closest<HTMLElement>('.dve-print-page');
		if (!view?.editable || !sheet) return;
		const scale = scaleOf(sheet);
		const start = boxOf(element);
		const origin = originOf(element);
		const lockAspect = !located.node.attrs.unsupported;
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
			const located = select(element);
			if (!located) return;
			event.preventDefault();
			event.stopImmediatePropagation();
			drag(event, located, element, grip?.dataset.floatHandle as FloatHandle | undefined);
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

	const observer = new MutationObserver(schedule);
	observer.observe(host.pages, { childList: true });
	observer.observe(host.pages, {
		attributes: true,
		attributeFilter: ['data-render-version'],
		subtree: true,
	});

	return {
		refresh: schedule,
		destroy() {
			controller.abort();
			observer.disconnect();
			if (redraw !== undefined) cancelAnimationFrame(redraw);
			clearFrame();
		},
	};
}
