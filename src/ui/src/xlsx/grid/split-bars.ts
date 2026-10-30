// The draggable split bars of View > Split. Dragging moves a bar live (a local override, no undo
// step per pixel) and commits one `setSplit` on release; a double click removes the bar, and both
// when it lands on the crossing.
import { clampSplitBar, pxToSplit, type SplitPane } from 'ooxml-core/xlsx';
import type { GridView } from './grid-view';
import { SPLIT_BAR } from './pane-chrome';

type Axis = 'x' | 'y';

export function installSplitBars(view: GridView): () => void {
	const commit = (next: SplitPane | undefined) => {
		const session = view.ctx.session();
		if (session && !view.ctx.readOnly()) session.setSplit(view.sheetIndex(), next);
	};

	const current = (): SplitPane | undefined => view.sheet()?.view.split;

	const remove = (axes: readonly Axis[]) => {
		const split = current();
		if (!split) return;
		const next: SplitPane = { ...split };
		for (const axis of axes) next[axis] = 0;
		commit(next.x > 0 || next.y > 0 ? next : undefined);
	};

	const install = (bar: HTMLElement, axis: Axis): (() => void) => {
		const other = axis === 'x' ? view.barH : view.barV;
		let pointer = -1;
		let moved = false;
		const position = (event: PointerEvent): number => {
			const { x, y } = view.clientToView(event.clientX, event.clientY);
			const g = view.geometry;
			return axis === 'x'
				? clampSplitBar(x - g.headerW, g.cellsW, view.metrics.zoom)
				: clampSplitBar(y - g.headerH, g.cellsH, view.metrics.zoom);
		};
		const down = (event: PointerEvent) => {
			if (event.button !== 0 || view.ctx.readOnly()) return;
			event.preventDefault();
			event.stopPropagation();
			pointer = event.pointerId;
			moved = false;
			bar.setPointerCapture?.(event.pointerId);
		};
		const move = (event: PointerEvent) => {
			if (event.pointerId !== pointer) return;
			moved = true;
			view.splitDrag = { ...view.splitDrag, [axis]: position(event) };
			view.paint();
		};
		const up = (event: PointerEvent) => {
			if (event.pointerId !== pointer) return;
			pointer = -1;
			bar.releasePointerCapture?.(event.pointerId);
			const drag = view.splitDrag;
			view.splitDrag = undefined;
			const split = current();
			if (!moved || !split || drag?.[axis] === undefined) return;
			commit({ ...split, [axis]: pxToSplit(drag[axis], view.metrics.zoom) });
		};
		const dbl = (event: MouseEvent) => {
			event.preventDefault();
			event.stopPropagation();
			if (view.ctx.readOnly()) return;
			const rect = other.getBoundingClientRect();
			// On the crossing: close to the other bar's centre line.
			const onCrossing =
				!other.hidden &&
				(axis === 'x'
					? Math.abs(event.clientY - (rect.top + rect.height / 2))
					: Math.abs(event.clientX - (rect.left + rect.width / 2))) <= SPLIT_BAR;
			remove(onCrossing ? ['x', 'y'] : [axis]);
		};
		bar.addEventListener('pointerdown', down);
		bar.addEventListener('pointermove', move);
		bar.addEventListener('pointerup', up);
		bar.addEventListener('pointercancel', up);
		bar.addEventListener('dblclick', dbl);
		return () => {
			bar.removeEventListener('pointerdown', down);
			bar.removeEventListener('pointermove', move);
			bar.removeEventListener('pointerup', up);
			bar.removeEventListener('pointercancel', up);
			bar.removeEventListener('dblclick', dbl);
		};
	};

	const offs = [install(view.barV, 'x'), install(view.barH, 'y')];
	return () => offs.forEach((off) => off());
}
