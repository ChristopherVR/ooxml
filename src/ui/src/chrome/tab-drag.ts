// Shared pointer drag for reordering a strip of tabs (Excel sheets, Visio pages). After a small
// movement threshold an insertion marker follows the pointer; releasing drops the tab into the
// marked gap, Escape or a cancelled pointer abandons the drag. Mouse, pen and touch all arrive as
// pointer events. The caller maps the gap to its own model and decides which tabs may move.

const THRESHOLD = 5;

/** The gap (0..n) nearest to `clientX`: before the first tab whose midpoint lies right of it. */
export function slotAt(tabs: HTMLElement[], clientX: number): number {
	for (let i = 0; i < tabs.length; i++) {
		const box = tabs[i]!.getBoundingClientRect();
		if (clientX < box.left + box.width / 2) return i;
	}
	return tabs.length;
}

/** The index a tab at `from` ends up at when dropped into gap `slot` of the same list. */
export function indexAfterDrop(from: number, slot: number): number {
	return slot > from ? slot - 1 : slot;
}

export interface TabDragHost {
	/** Receives the pointer events and the pointer capture. */
	area: HTMLElement;
	/** The tab elements in display order. */
	tabs(): HTMLElement[];
	/** Whether this pointer-down target may start a drag. */
	canDrag(tab: HTMLElement): boolean;
	/** Class names for the dragged tab and the insertion marker. */
	draggingClass: string;
	markerClass: string;
	/** Where the marker is appended, and the element its `left` is measured from. */
	markerParent: ParentNode;
	markerBox: HTMLElement;
	/** Selector matching a tab element. */
	tabSelector: string;
	/** Inclusive gap range the tab may be dropped into; defaults to every gap. */
	limit?(tab: HTMLElement): [number, number];
	/** Called once with the tab and the chosen gap when a drag completes. */
	drop(tab: HTMLElement, slot: number): void;
}

interface DragState {
	tab: HTMLElement;
	x: number;
	pointer: number;
	dragging: boolean;
	slot: number;
}

/** Wires pointer dragging on `host.area`; returns a disposer. */
export function wireTabDrag(host: TabDragHost): () => void {
	const { area } = host;
	const doc = area.ownerDocument;
	let state: DragState | undefined;
	let marker: HTMLElement | undefined;

	const stop = (event: Event) => event.stopPropagation();
	// Swallow the click that follows a drag so it does not activate a tab.
	const swallowClick = () => {
		area.addEventListener('click', stop, { capture: true, once: true });
		setTimeout(() => area.removeEventListener('click', stop, true), 0);
	};
	const clear = () => {
		marker?.remove();
		marker = undefined;
		state?.tab.classList.remove(host.draggingClass);
		state = undefined;
		doc.removeEventListener('keydown', onKey, true);
	};
	const onKey = (event: KeyboardEvent) => {
		if (event.key !== 'Escape' || !state) return;
		event.preventDefault();
		event.stopPropagation();
		if (state.dragging) swallowClick();
		clear();
	};
	const onDown = (event: PointerEvent) => {
		if (event.button !== 0) return;
		const tab = (event.target as Element).closest?.<HTMLElement>(host.tabSelector);
		if (!tab || !host.canDrag(tab)) return;
		state = { tab, x: event.clientX, pointer: event.pointerId, dragging: false, slot: -1 };
		doc.addEventListener('keydown', onKey, true);
	};
	const onMove = (event: PointerEvent) => {
		if (!state || event.pointerId !== state.pointer) return;
		if (!state.dragging) {
			if (Math.abs(event.clientX - state.x) < THRESHOLD) return;
			state.dragging = true;
			state.tab.classList.add(host.draggingClass);
			try {
				area.setPointerCapture?.(event.pointerId);
			} catch {
				// Synthetic pointers (tests) cannot be captured.
			}
		}
		const list = host.tabs();
		const [lo, hi] = host.limit?.(state.tab) ?? [0, list.length];
		state.slot = Math.max(lo, Math.min(hi, slotAt(list, event.clientX)));
		marker ??= doc.createElement('div');
		marker.className = host.markerClass;
		marker.setAttribute('aria-hidden', 'true');
		const box = host.markerBox.getBoundingClientRect();
		const ref = list[state.slot] ?? list[list.length - 1];
		const rect = ref?.getBoundingClientRect();
		const x = !rect ? 0 : state.slot < list.length ? rect.left : rect.right;
		marker.style.left = `${x - box.left + host.markerBox.scrollLeft}px`;
		host.markerParent.append(marker);
	};
	const onUp = (event: PointerEvent) => {
		if (!state || event.pointerId !== state.pointer) return;
		const { dragging, tab, slot } = state;
		clear();
		if (!dragging || slot < 0) return;
		swallowClick();
		host.drop(tab, slot);
	};
	const onCancel = (event: PointerEvent) => {
		if (state && event.pointerId === state.pointer) clear();
	};
	area.addEventListener('pointerdown', onDown);
	area.addEventListener('pointermove', onMove);
	area.addEventListener('pointerup', onUp);
	area.addEventListener('pointercancel', onCancel);
	return () => {
		clear();
		area.removeEventListener('pointerdown', onDown);
		area.removeEventListener('pointermove', onMove);
		area.removeEventListener('pointerup', onUp);
		area.removeEventListener('pointercancel', onCancel);
	};
}
