// Drag a tab to reorder sheets: after a 5px threshold an insertion marker follows the pointer;
// dropping moves the sheet with the core's `moveSheet` and keeps it active.
import { h } from '../grid/dom.js';

const THRESHOLD = 5;

/** The sheet index a sheet dropped before visible slot `slot` ends up at (`moveSheet` target). */
export function dropTarget(visible: number[], from: number, slot: number): number {
	if (!visible.length) return from;
	const last = visible[visible.length - 1] ?? from;
	const before = slot < visible.length ? (visible[slot] ?? last + 1) : last + 1;
	return before > from ? before - 1 : before;
}

/** The visible slot (0..n) whose gap is nearest to `clientX`. */
export function slotAt(tabs: HTMLElement[], clientX: number): number {
	for (let i = 0; i < tabs.length; i++) {
		const box = tabs[i]!.getBoundingClientRect();
		if (clientX < box.left + box.width / 2) return i;
	}
	return tabs.length;
}

export interface DragHost {
	strip: HTMLElement;
	enabled(): boolean;
	visible(): number[];
	move(from: number, to: number): void;
}

/** Wires pointer dragging on the strip; returns a disposer. */
export function wireTabDrag(host: DragHost): () => void {
	const { strip } = host;
	let state:
		| {
				from: number;
				tab: HTMLElement;
				x: number;
				pointer: number;
				dragging: boolean;
				slot: number;
		  }
		| undefined;
	let marker: HTMLElement | undefined;

	const tabs = () => [...strip.querySelectorAll<HTMLElement>('.xst-tab')];
	const clear = () => {
		marker?.remove();
		marker = undefined;
		state?.tab.classList.remove('xst-dragging');
		state = undefined;
	};
	const onDown = (event: PointerEvent) => {
		if (event.button !== 0 || !host.enabled()) return;
		const tab = (event.target as Element).closest?.<HTMLElement>('.xst-tab');
		if (!tab || tab.dataset.sheet === undefined) return;
		state = {
			from: Number(tab.dataset.sheet),
			tab,
			x: event.clientX,
			pointer: event.pointerId,
			dragging: false,
			slot: -1,
		};
	};
	const onMove = (event: PointerEvent) => {
		if (!state || event.pointerId !== state.pointer) return;
		if (!state.dragging) {
			if (Math.abs(event.clientX - state.x) < THRESHOLD) return;
			state.dragging = true;
			state.tab.classList.add('xst-dragging');
			try {
				strip.setPointerCapture?.(event.pointerId);
			} catch {
				// Synthetic pointers (tests) cannot be captured.
			}
		}
		const list = tabs();
		state.slot = slotAt(list, event.clientX);
		marker ??= h(strip.ownerDocument, 'div', 'xst-marker', { 'aria-hidden': 'true' });
		const stripBox = strip.getBoundingClientRect();
		const ref = list[state.slot] ?? list[list.length - 1];
		const box = ref?.getBoundingClientRect();
		const x = !box ? 0 : state.slot < list.length ? box.left : box.right;
		marker.style.left = `${x - stripBox.left + strip.scrollLeft}px`;
		strip.append(marker);
	};
	const onUp = (event: PointerEvent) => {
		if (!state || event.pointerId !== state.pointer) return;
		const { dragging, from, slot } = state;
		clear();
		if (!dragging || slot < 0) return;
		// Swallow the click that follows the drag so it does not re-activate another tab.
		strip.addEventListener('click', stop, { capture: true, once: true });
		setTimeout(() => strip.removeEventListener('click', stop, true), 0);
		const to = dropTarget(host.visible(), from, slot);
		if (to !== from) host.move(from, to);
	};
	const stop = (e: Event) => e.stopPropagation();
	const onCancel = () => clear();
	strip.addEventListener('pointerdown', onDown);
	strip.addEventListener('pointermove', onMove);
	strip.addEventListener('pointerup', onUp);
	strip.addEventListener('pointercancel', onCancel);
	return () => {
		clear();
		strip.removeEventListener('pointerdown', onDown);
		strip.removeEventListener('pointermove', onMove);
		strip.removeEventListener('pointerup', onUp);
		strip.removeEventListener('pointercancel', onCancel);
	};
}
