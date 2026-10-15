// Page Break Preview: paints the print area, the shaded cells outside it, "Page N" watermarks and
// the blue page-break lines over the grid, and lets a drag move a break or an edge of the print
// area. Every decision (where the breaks are, what a drag changes) comes from the core's
// `pageBreakPreview`, `planBreakDrag` and `resizePrintArea`; this file only paints and routes.
import {
	pageBreakPreview,
	planBreakDrag,
	resizePrintArea,
	type PageBreakLine,
	type PageBreakPreview,
	type PrintAreaEdge,
} from 'ooxml-core/xlsx';
import { h, place } from './dom';
import type { GridView } from './grid-view';

export const PAGE_BREAK_CSS = `
.xgpb-layer{position:absolute;inset:0;pointer-events:none;z-index:3;overflow:hidden}
.xgpb-layer[hidden]{display:none}
.xgpb-shade{position:absolute;background:rgba(128,128,128,.32)}
.xgpb-area{position:absolute;box-sizing:border-box;border:2px solid #2b579a}
.xgpb-mark{position:absolute;display:flex;align-items:center;justify-content:center;overflow:hidden;
	font:700 clamp(28px,6vw,72px)/1 "Segoe UI",system-ui,sans-serif;color:rgba(128,128,128,.42);white-space:nowrap}
.xgpb-line{position:absolute;pointer-events:none;background:#2b579a}
.xgpb-line.xgpb-auto{background:none}
.xgpb-line.xgpb-auto.xgpb-row{border-top:2px dashed #2b579a}
.xgpb-line.xgpb-auto.xgpb-col{border-left:2px dashed #2b579a}
.xgpb-grab{position:absolute;pointer-events:auto;z-index:1;touch-action:none}
.xgpb-grab.xgpb-row{cursor:ns-resize}
.xgpb-grab.xgpb-col{cursor:ew-resize}
.xgpb-ghost{position:absolute;pointer-events:none;background:rgba(43,87,154,.55);z-index:2}
`;

type Grab = { kind: 'line'; line: PageBreakLine } | { kind: 'edge'; edge: PrintAreaEdge };
const GRAB = 8;

export function installPageBreakPreview(view: GridView): () => void {
	const doc = view.doc;
	const layer = h(doc, 'div', 'xgpb-layer', { 'data-part': 'page-break-preview' });
	layer.hidden = true;
	const ghost = h(doc, 'div', 'xgpb-ghost');
	ghost.hidden = true;
	view.view.append(layer);
	let cache: { key: unknown; sheet: number; preview: PageBreakPreview | undefined } | undefined;
	let dragging = false;

	const active = (): boolean => view.sheet()?.view.mode === 'pageBreakPreview';
	const preview = (): PageBreakPreview | undefined => {
		const workbook = view.workbook();
		const sheet = view.sheetIndex();
		if (!workbook) return undefined;
		if (cache?.key === workbook && cache.sheet === sheet) return cache.preview;
		let result: PageBreakPreview | undefined;
		try {
			result = pageBreakPreview(workbook, sheet);
		} catch {
			result = undefined;
		}
		cache = { key: workbook, sheet, preview: result };
		return result;
	};
	const offModel = view.ctx.onModelChange(() => (cache = undefined));

	const target = (grab: Grab, clientX: number, clientY: number): number => {
		const g = view.geometry;
		const { x, y } = view.clientToView(clientX, clientY);
		const horizontal =
			grab.kind === 'line'
				? grab.line.axis === 'col'
				: grab.edge === 'left' || grab.edge === 'right';
		if (horizontal) {
			const col = g.colAtX(x);
			return x > g.colX(col) + view.metrics.colWidth(col) / 2 ? col + 1 : col;
		}
		const row = g.rowAtY(y);
		return y > g.rowY(row) + view.metrics.rowHeight(row) / 2 ? row + 1 : row;
	};

	const apply = (grab: Grab, boundary: number) => {
		const session = view.ctx.session();
		const data = preview();
		if (!session || !data) return;
		const sheet = view.sheetIndex();
		if (grab.kind === 'line') {
			const edits = planBreakDrag(data, grab.line, boundary);
			if (!edits.length) {
				if (data.fit)
					view.ctx.toast(
						view.ctx.t('Page breaks are set automatically while Fit to pages is on.'),
						'info',
					);
				return;
			}
			session.batch('Move page break', () => {
				for (const edit of edits) session.setPageBreak(sheet, edit.axis, edit.before, edit.present);
			});
			return;
		}
		if (!data.area) return;
		const last = grab.edge === 'bottom' || grab.edge === 'right';
		const next = resizePrintArea(data.area, grab.edge, last ? boundary - 1 : boundary);
		if (next) session.setPageSetup(sheet, { printArea: next });
	};

	const grabHandle = (
		grab: Grab,
		x: number,
		y: number,
		w: number,
		h2: number,
		axis: 'row' | 'col',
	) => {
		const node = h(doc, 'div', `xgpb-grab xgpb-${axis}`);
		place(node, x, y, w, h2);
		node.addEventListener('pointerdown', (event) => {
			if (event.button !== 0 || view.ctx.readOnly()) return;
			event.preventDefault();
			event.stopPropagation();
			node.setPointerCapture?.(event.pointerId);
			dragging = true;
			ghost.hidden = false;
			layer.append(ghost);
			let boundary = -1;
			const move = (e: PointerEvent) => {
				boundary = target(grab, e.clientX, e.clientY);
				const g = view.geometry;
				if (axis === 'row') place(ghost, g.headerW, g.rowY(boundary) - 1, g.cellsW, 3);
				else place(ghost, g.colX(boundary) - 1, g.headerH, 3, g.cellsH);
			};
			const finish = (e: PointerEvent, commit: boolean) => {
				node.removeEventListener('pointermove', move);
				node.removeEventListener('pointerup', up);
				node.removeEventListener('pointercancel', cancel);
				dragging = false;
				ghost.hidden = true;
				if (commit && boundary >= 0) apply(grab, boundary);
				view.schedule();
				void e;
			};
			const up = (e: PointerEvent) => finish(e, true);
			const cancel = (e: PointerEvent) => finish(e, false);
			node.addEventListener('pointermove', move);
			node.addEventListener('pointerup', up);
			node.addEventListener('pointercancel', cancel);
		});
		return node;
	};

	const hook = () => {
		if (dragging) return;
		const data = active() ? preview() : undefined;
		layer.hidden = !data;
		if (!data) {
			layer.replaceChildren();
			return;
		}
		const g = view.geometry;
		const out: HTMLElement[] = [];
		const x0 = g.headerW;
		const y0 = g.headerH;
		const x1 = g.width;
		const y1 = g.height;
		const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v));
		if (data.area) {
			const box = g.rangeBox(data.area);
			const ax0 = clamp(box.x, x0, x1);
			const ay0 = clamp(box.y, y0, y1);
			const ax1 = clamp(box.x + box.w, x0, x1);
			const ay1 = clamp(box.y + box.h, y0, y1);
			for (const [x, y, w, hh] of [
				[x0, y0, x1 - x0, ay0 - y0],
				[x0, ay1, x1 - x0, y1 - ay1],
				[x0, ay0, ax0 - x0, ay1 - ay0],
				[ax1, ay0, x1 - ax1, ay1 - ay0],
			] as const) {
				if (w <= 0 || hh <= 0) continue;
				const shade = h(doc, 'div', 'xgpb-shade');
				place(shade, x, y, w, hh);
				out.push(shade);
			}
			const border = h(doc, 'div', 'xgpb-area');
			place(border, box.x - 1, box.y - 1, box.w + 2, box.h + 2);
			out.push(border);
			const edges: [PrintAreaEdge, number, number, number, number, 'row' | 'col'][] = [
				['top', box.x, box.y - GRAB / 2, box.w, GRAB, 'row'],
				['bottom', box.x, box.y + box.h - GRAB / 2, box.w, GRAB, 'row'],
				['left', box.x - GRAB / 2, box.y, GRAB, box.h, 'col'],
				['right', box.x + box.w - GRAB / 2, box.y, GRAB, box.h, 'col'],
			];
			for (const [edge, x, y, w, hh, axis] of edges) {
				if (x + w < x0 || y + hh < y0 || x > x1 || y > y1) continue;
				out.push(grabHandle({ kind: 'edge', edge }, x, y, w, hh, axis));
			}
		}
		for (const page of data.pages) {
			const box = g.rangeBox(page.range);
			const mark = h(doc, 'div', 'xgpb-mark');
			const mx = clamp(box.x, x0, x1);
			const my = clamp(box.y, y0, y1);
			place(mark, mx, my, clamp(box.x + box.w, x0, x1) - mx, clamp(box.y + box.h, y0, y1) - my);
			mark.textContent = view.ctx.t('Page {number}', { number: page.number });
			out.push(mark);
		}
		for (const line of data.lines) {
			const row = line.axis === 'row';
			const where = row ? g.rowY(line.before) : g.colX(line.before);
			if (where < (row ? y0 : x0) || where > (row ? y1 : x1)) continue;
			const area = data.area ? g.rangeBox(data.area) : undefined;
			const from = row ? (area?.x ?? x0) : (area?.y ?? y0);
			const length = row ? (area?.w ?? x1 - x0) : (area?.h ?? y1 - y0);
			const mark = h(
				doc,
				'div',
				`xgpb-line xgpb-${line.axis} ${line.manual ? 'xgpb-manual' : 'xgpb-auto'}`,
			);
			if (row) place(mark, from, where - 1, length, line.manual ? 3 : 0);
			else place(mark, where - 1, from, line.manual ? 3 : 0, length);
			out.push(mark);
			out.push(
				row
					? grabHandle({ kind: 'line', line }, from, where - GRAB / 2, length, GRAB, 'row')
					: grabHandle({ kind: 'line', line }, where - GRAB / 2, from, GRAB, length, 'col'),
			);
		}
		layer.replaceChildren(...out);
	};
	view.hooks.add(hook);
	return () => {
		view.hooks.delete(hook);
		offModel();
		layer.remove();
	};
}
