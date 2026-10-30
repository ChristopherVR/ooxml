// Drag a tab to reorder sheets: after a 5px threshold an insertion marker follows the pointer;
// dropping moves the sheet with the core's `moveSheet` and keeps it active. The pointer handling is
// the shared tab drag in `chrome/tab-drag`.
import { wireTabDrag as wireSharedTabDrag } from '../../chrome/tab-drag';

export { slotAt } from '../../chrome/tab-drag';

/** The sheet index a sheet dropped before visible slot `slot` ends up at (`moveSheet` target). */
export function dropTarget(visible: number[], from: number, slot: number): number {
	if (!visible.length) return from;
	const last = visible[visible.length - 1] ?? from;
	const before = slot < visible.length ? (visible[slot] ?? last + 1) : last + 1;
	return before > from ? before - 1 : before;
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
	return wireSharedTabDrag({
		area: strip,
		tabs: () => [...strip.querySelectorAll<HTMLElement>('.xst-tab')],
		tabSelector: '.xst-tab',
		canDrag: (tab) => host.enabled() && tab.dataset.sheet !== undefined,
		draggingClass: 'xst-dragging',
		markerClass: 'xst-marker',
		markerParent: strip,
		markerBox: strip,
		drop(tab, slot) {
			const from = Number(tab.dataset.sheet);
			const to = dropTarget(host.visible(), from, slot);
			if (to !== from) host.move(from, to);
		},
	});
}
