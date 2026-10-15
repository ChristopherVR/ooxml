// Page Layout view: the sheet drawn as pages (paper size, orientation and margins from the page
// setup) with rulers, per-page headings, header and footer bands and cell editing. The page plan
// and every hit test come from the core's `pageViewLayout` and `pageViewHit`; this file paints
// them and routes pointer and key input to the edit session.
import {
	cellInputText,
	createGridMetrics,
	formatAddress,
	pageOfCell,
	pageViewCellRect,
	pageViewHit,
	pageViewLayout,
	type GridMetrics,
	type PageViewLayout,
	type PageViewPage,
	type PageViewRect,
} from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { extendSelection, selectCell } from 'ooxml-core/xlsx/ui';
import { pageTable } from '../print-pages';
import { editBand, renderBand, type BandOptions } from './page-layout-bands';
import { PAGE_LAYOUT_CSS } from './page-layout-css';
import { paintRulers } from './page-layout-rulers';
import { createPageEditor } from './page-layout-editor';

const PAD = 24;
const el = <K extends keyof HTMLElementTagNameMap>(doc: Document, tag: K, cls: string) => {
	const node = doc.createElement(tag);
	node.className = cls;
	return node;
};
const css = (node: HTMLElement, rect: PageViewRect, origin = { x: 0, y: 0 }) => {
	node.style.cssText += `;left:${rect.x - origin.x}px;top:${rect.y - origin.y}px;width:${rect.width}px;height:${rect.height}px`;
};

export interface PageLayoutView {
	readonly element: HTMLElement;
	refresh(): void;
	setActive(on: boolean): void;
	destroy(): void;
}

export function createPageLayoutView(
	ctx: EditorContext,
	host: HTMLElement,
	fileName: () => string,
): PageLayoutView {
	const doc = host.ownerDocument;
	if (!ctx.root.querySelector('style[data-xpl]')) {
		const style = doc.createElement('style');
		style.dataset.xpl = '';
		style.textContent = PAGE_LAYOUT_CSS;
		ctx.root.append(style);
	}
	const root = el(doc, 'div', 'xpl-root');
	root.tabIndex = 0;
	root.hidden = true;
	root.setAttribute('part', 'page-layout');
	root.setAttribute('role', 'application');
	root.setAttribute('aria-label', ctx.t('Page Layout view'));
	const sizer = el(doc, 'div', 'xpl-sizer');
	const corner = el(doc, 'div', 'xpl-corner');
	const rulerH = el(doc, 'div', 'xpl-ruler-h');
	const rulerV = el(doc, 'div', 'xpl-ruler-v');
	const canvas = el(doc, 'div', 'xpl-canvas');
	const scaled = el(doc, 'div', 'xpl-scaled');
	const selection = el(doc, 'div', 'xpl-sel');
	canvas.append(scaled);
	sizer.append(corner, rulerH, rulerV, canvas);
	root.append(sizer);
	host.append(root);

	let active = false;
	let layout: PageViewLayout | undefined;
	let metrics: GridMetrics | undefined;
	let bandEditing = false;
	let dragFrom: { row: number; col: number } | undefined;
	let frame = 0;
	let renderedKey = '';
	const filled = new Set<number>();
	const pageNodes = new Map<number, HTMLElement>();

	const zoom = () =>
		Math.max(
			0.1,
			(ctx.grid()?.zoom() ?? ctx.workbook()?.sheets[ctx.activeSheet()]?.view.zoom ?? 100) / 100,
		);
	const sheetIndex = () => ctx.activeSheet();
	const bandOptions = (): BandOptions | undefined =>
		layout && ctx.workbook()
			? { ctx, doc, layout, sheetIndex: sheetIndex(), fileName: fileName(), date: new Date() }
			: undefined;

	const fillPage = (page: PageViewPage) => {
		const node = pageNodes.get(page.number);
		const workbook = ctx.workbook();
		if (!node || !workbook || !layout || filled.has(page.number)) return;
		filled.add(page.number);
		const sheet = workbook.sheets[sheetIndex()]!;
		const content = node.querySelector<HTMLElement>('.xpl-content')!;
		if (page.blank) {
			const blank = el(doc, 'div', 'xpl-blank');
			blank.textContent = ctx.t('Click to add data');
			content.append(blank);
			return;
		}
		const inner = el(doc, 'div', 'xpl-inner');
		inner.style.transform = `scale(${layout.scale})`;
		inner.innerHTML = pageTable(
			workbook,
			sheetIndex(),
			{
				rows: page.rows,
				columns: page.columns,
				width: page.colLefts.length
					? page.colLefts[page.colLefts.length - 1]! +
						metrics!.colWidth(page.columns[page.columns.length - 1]!)
					: 0,
			},
			{ headings: false, gridLines: sheet.view.showGridLines !== false },
		);
		content.append(inner);
	};

	const fillVisible = () => {
		if (!active || !layout) return;
		const z = zoom();
		const top = root.scrollTop / z - PAD;
		const bottom = (root.scrollTop + root.clientHeight) / z;
		for (const page of layout.pages)
			if (page.paper.y + page.paper.height > top - 600 && page.paper.y < bottom + 600)
				fillPage(page);
	};

	const headings = (page: PageViewPage, node: HTMLElement) => {
		const sheet = ctx.workbook()?.sheets[sheetIndex()];
		if (!layout || !metrics || page.blank || sheet?.view.showHeaders === false) return;
		page.columns.forEach((col, i) => {
			const cell = el(doc, 'div', 'xpl-colhdr');
			cell.textContent = formatAddress({ row: 0, col }).replace(/\d+$/, '');
			cell.style.cssText = `left:${page.content.x - page.paper.x + page.colLefts[i]! * layout!.scale}px;top:${page.content.y - page.paper.y - 17}px;width:${metrics!.colWidth(col) * layout!.scale}px;height:16px`;
			node.append(cell);
		});
		page.rows.forEach((row, i) => {
			const cell = el(doc, 'div', 'xpl-rowhdr');
			cell.textContent = String(row + 1);
			cell.style.cssText = `left:${page.content.x - page.paper.x - 34}px;top:${page.content.y - page.paper.y + page.rowTops[i]! * layout!.scale}px;width:32px;height:${metrics!.rowHeight(row) * layout!.scale}px`;
			node.append(cell);
		});
	};

	const render = () => {
		frame = 0;
		if (!active) return;
		const workbook = ctx.workbook();
		scaled.replaceChildren();
		pageNodes.clear();
		filled.clear();
		layout = undefined;
		if (!workbook || !workbook.sheets[sheetIndex()]) return;
		try {
			layout = pageViewLayout(workbook, sheetIndex());
		} catch (error) {
			const message = el(doc, 'div', 'xpl-error');
			message.textContent = ctx.t(
				error instanceof Error ? error.message : 'Page Layout view is not available.',
			);
			scaled.append(message);
			return;
		}
		metrics = createGridMetrics(workbook.sheets[sheetIndex()]!, { zoom: 100 });
		const z = zoom();
		renderedKey = `${z * 100}|${ctx.locale()}`;
		const width = (layout.width + PAD * 2) * z;
		const height = (layout.height + PAD * 2) * z;
		canvas.style.cssText = `width:${width}px;height:${height}px`;
		scaled.style.cssText = `width:${layout.width + PAD * 2}px;height:${layout.height + PAD * 2}px;transform:scale(${z})`;
		paintRulers(doc, rulerH, rulerV, layout, z, PAD);
		const bands = bandOptions()!;
		for (const page of layout.pages) {
			const node = el(doc, 'div', 'xpl-page');
			node.dataset.page = String(page.number);
			css(node, { ...page.paper, x: page.paper.x + PAD, y: page.paper.y + PAD });
			node.append(renderBand(bands, page, 'header'), renderBand(bands, page, 'footer'));
			const content = el(doc, 'div', 'xpl-content');
			css(content, page.content, page.paper);
			node.append(content);
			headings(page, node);
			scaled.append(node);
			pageNodes.set(page.number, node);
		}
		selection.style.cssText = `left:${PAD}px;top:${PAD}px`;
		scaled.append(selection);
		fillVisible();
		paintSelection();
	};

	const schedule = () => {
		if (!frame) frame = (doc.defaultView ?? window).requestAnimationFrame(render);
	};

	const toCanvas = (clientX: number, clientY: number) => {
		const rect = scaled.getBoundingClientRect();
		const z = zoom();
		return { x: (clientX - rect.left) / z - PAD, y: (clientY - rect.top) / z - PAD };
	};

	const editor = createPageEditor({
		ctx,
		doc,
		scaled,
		rectOf: (row, col) => {
			if (!layout || !metrics) return undefined;
			const n = pageOfCell(layout, row, col);
			const page = layout.pages.find((p) => p.number === (n ?? 1));
			if (page?.blank)
				return {
					...page.content,
					x: page.content.x + PAD,
					y: page.content.y + PAD,
					width: 64,
					height: 20,
				};
			const rect = n === undefined ? undefined : pageViewCellRect(layout, metrics, n, row, col);
			return rect && { ...rect, x: rect.x + PAD, y: rect.y + PAD };
		},
		scale: () => layout?.scale ?? 1,
		done: () => root.focus({ preventScroll: true }),
	});

	function paintSelection(): void {
		selection.replaceChildren();
		const sel = ctx.selection.get();
		if (!layout || !metrics || sel.sheet !== sheetIndex()) return;
		for (const page of layout.pages) {
			if (page.blank) continue;
			for (const range of sel.ranges) {
				const rows = page.rows.filter((r) => r >= range.start.row && r <= range.end.row);
				const cols = page.columns.filter((c) => c >= range.start.col && c <= range.end.col);
				if (!rows.length || !cols.length) continue;
				const a = pageViewCellRect(layout, metrics, page.number, rows[0]!, cols[0]!);
				const b = pageViewCellRect(layout, metrics, page.number, rows.at(-1)!, cols.at(-1)!);
				if (!a || !b) continue;
				const shade = el(doc, 'div', 'xpl-shade');
				css(shade, { x: a.x, y: a.y, width: b.x + b.width - a.x, height: b.y + b.height - a.y });
				selection.append(shade);
			}
		}
		const at = sel.active;
		const n = pageOfCell(layout, at.row, at.col);
		const rect = n === undefined ? undefined : pageViewCellRect(layout, metrics, n, at.row, at.col);
		if (rect) {
			const box = el(doc, 'div', 'xpl-active');
			css(box, rect);
			selection.append(box);
		}
	}

	const select = (row: number, col: number, extend: boolean) => {
		const sheet = ctx.workbook()?.sheets[sheetIndex()];
		const current = ctx.selection.get();
		ctx.selection.set(
			extend && current.sheet === sheetIndex()
				? extendSelection(current, sheet, { row, col })
				: selectCell(sheetIndex(), sheet, { row, col }),
		);
	};

	const openBand = (
		page: PageViewPage,
		kind: 'header' | 'footer',
		section: 'left' | 'center' | 'right',
	) => {
		const o = bandOptions();
		const node = pageNodes.get(page.number)?.querySelector<HTMLElement>(`.xpl-${kind}`);
		if (!o || !node || ctx.readOnly()) return;
		bandEditing = true;
		editBand(o, page, node, kind, section, () => {
			bandEditing = false;
			render();
			root.focus({ preventScroll: true });
		});
	};

	root.addEventListener('pointerdown', (event) => {
		if (event.button !== 0 || !layout || !metrics || editor.editing) return;
		const target = event.target as HTMLElement;
		if (target.closest('.xpl-ruler-h,.xpl-ruler-v,.xpl-corner') || bandEditing) return;
		const { x, y } = toCanvas(event.clientX, event.clientY);
		const hit = pageViewHit(layout, metrics, x, y);
		if (!hit) return;
		const page = layout.pages.find((p) => p.number === hit.page)!;
		if (hit.kind === 'header' || hit.kind === 'footer') {
			event.preventDefault();
			openBand(page, hit.kind, hit.section);
		} else if (hit.kind === 'cell') {
			root.focus({ preventScroll: true });
			event.preventDefault();
			select(hit.row, hit.col, event.shiftKey);
			dragFrom = { row: hit.row, col: hit.col };
		} else if (hit.kind === 'blank') {
			event.preventDefault();
			select(0, 0, false);
			editor.begin(0, 0, '');
		}
	});
	root.addEventListener('pointermove', (event) => {
		if (!dragFrom || !layout || !metrics || event.buttons !== 1) return;
		const { x, y } = toCanvas(event.clientX, event.clientY);
		const hit = pageViewHit(layout, metrics, x, y);
		if (hit?.kind === 'cell') {
			const sheet = ctx.workbook()?.sheets[sheetIndex()];
			const base = selectCell(sheetIndex(), sheet, dragFrom);
			ctx.selection.set(extendSelection(base, sheet, { row: hit.row, col: hit.col }));
		}
	});
	const release = () => (dragFrom = undefined);
	root.addEventListener('pointerup', release);
	root.addEventListener('pointercancel', release);
	root.addEventListener('dblclick', (event) => {
		if (!layout || !metrics || editor.editing) return;
		const { x, y } = toCanvas(event.clientX, event.clientY);
		const hit = pageViewHit(layout, metrics, x, y);
		const workbook = ctx.workbook();
		if (hit?.kind === 'cell' && workbook)
			editor.begin(hit.row, hit.col, cellInputText(workbook, sheetIndex(), hit.row, hit.col));
	});
	root.addEventListener('scroll', () => fillVisible(), { passive: true });
	root.addEventListener('keydown', (event) => {
		if (editor.editing || bandEditing || event.target !== root) return;
		const sel = ctx.selection.get();
		const at = sel.active;
		const workbook = ctx.workbook();
		const move = (dr: number, dc: number) => {
			event.preventDefault();
			select(Math.max(0, at.row + dr), Math.max(0, at.col + dc), event.shiftKey);
			const rect = editor.rectFor(ctx.selection.get().active);
			if (rect) root.scrollTo({ top: Math.max(0, (rect.y - 120) * zoom() - 40) });
		};
		if (event.key === 'ArrowDown') move(1, 0);
		else if (event.key === 'ArrowUp') move(-1, 0);
		else if (event.key === 'ArrowLeft') move(0, -1);
		else if (event.key === 'ArrowRight') move(0, 1);
		else if (event.key === 'F2' || event.key === 'Enter') {
			event.preventDefault();
			if (workbook)
				editor.begin(at.row, at.col, cellInputText(workbook, sheetIndex(), at.row, at.col));
		} else if (event.key === 'Delete' || event.key === 'Backspace') {
			event.preventDefault();
			const session = ctx.session();
			if (session && !ctx.readOnly())
				session.batch('Clear contents', () => {
					for (const range of sel.ranges)
						for (let r = range.start.row; r <= range.end.row; r++)
							for (let c = range.start.col; c <= range.end.col; c++)
								session.setCellInput(sheetIndex(), r, c, '');
				});
		} else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
			event.preventDefault();
			editor.begin(at.row, at.col, event.key);
		}
	});

	const offSelection = ctx.selection.onChange(() => active && !frame && paintSelection());
	const offModel = ctx.onModelChange(() => active && !editor.editing && !bandEditing && schedule());

	return {
		element: root,
		refresh() {
			if (active && renderedKey !== `${zoom()}|${ctx.locale()}`) schedule();
		},
		setActive(on) {
			if (active === on) return;
			active = on;
			root.hidden = !on;
			if (on) render();
			else {
				editor.cancel();
				scaled.replaceChildren();
			}
		},
		destroy() {
			offSelection();
			offModel();
			root.remove();
		},
	};
}
