// The sheet tab bar under the grid: scroll arrows, "+" new sheet and one tab per visible sheet
// (activate, rename by double-click, drag to reorder, context menu).
import { nextSheetName } from '@christophervr/xlsx-core';
import { sheetBaseName } from '../localization.js';
import type { EditorContext } from '../context.js';
import { openContextMenu, tabMenu } from '../context-menu/index.js';
import { ensureStyle, h } from '../grid/dom.js';
import { wireTabDrag } from './drag.js';
import { startRename, type RenameHandle } from './rename.js';
import { renderTabs, sheetOf, tabFor, visibleSheetIndices } from './render.js';
import { TABS_CSS } from './styles.js';

export { dropTarget, slotAt } from './drag.js';
export { visibleSheetIndices } from './render.js';

export function mountSheetTabs(ctx: EditorContext, container: HTMLElement): () => void {
	const doc = container.ownerDocument;
	ensureStyle(ctx.root, 'xst', TABS_CSS);
	const root = h(doc, 'div', 'xst', { part: 'sheet-tabs' });
	const nav = h(doc, 'div', 'xst-nav');
	const button = (cls: string, text: string, label: string) => {
		const node = h(doc, 'button', `xst-btn ${cls}`, {
			type: 'button',
			'aria-label': ctx.t(label),
			title: ctx.t(label),
		});
		node.textContent = text;
		return node;
	};
	const prev = button('xst-prev', '◂', 'Scroll to the previous sheet');
	const next = button('xst-next', '▸', 'Scroll to the next sheet');
	const add = button('xst-add', '+', 'New sheet');
	nav.append(prev, next, add);
	const strip = h(doc, 'div', 'xst-strip', { role: 'tablist', 'aria-label': ctx.t('Sheets') });
	root.append(nav, strip);
	container.append(root);

	let focusable = ctx.activeSheet();
	let renaming: RenameHandle | undefined;
	const editable = () => {
		const session = ctx.session();
		return Boolean(session) && !ctx.readOnly() && !session?.workbook.structureLocked;
	};

	const scrollIntoView = (sheet: number) => {
		const tab = tabFor(strip, sheet);
		if (!tab) return;
		const left = tab.offsetLeft;
		const right = left + tab.offsetWidth;
		if (left < strip.scrollLeft) strip.scrollLeft = left;
		else if (right > strip.scrollLeft + strip.clientWidth)
			strip.scrollLeft = right - strip.clientWidth;
	};

	const render = () => {
		if (renaming) return;
		const workbook = ctx.workbook();
		const active = ctx.activeSheet();
		add.hidden = !editable();
		if (!workbook) {
			strip.replaceChildren();
			return;
		}
		const visible = visibleSheetIndices(workbook);
		if (!visible.includes(focusable))
			focusable = visible.includes(active) ? active : (visible[0] ?? 0);
		renderTabs(doc, strip, { workbook, active, focusable });
		scrollIntoView(active);
	};

	const activate = (sheet: number) => {
		focusable = sheet;
		if (ctx.activeSheet() !== sheet) ctx.setActiveSheet(sheet);
		render();
	};

	const rename = (sheet: number) => {
		if (!editable()) return;
		activate(sheet);
		const tab = tabFor(strip, sheet);
		if (!tab) return;
		renaming = startRename(ctx, tab, sheet, () => {
			renaming = undefined;
			render();
			tabFor(strip, sheet)?.focus();
		});
	};

	const openMenu = (sheet: number, x: number, y: number) => {
		activate(sheet);
		const workbook = ctx.workbook();
		if (!workbook) return;
		const visible = visibleSheetIndices(workbook).length;
		const entries = tabMenu(
			{
				readOnly: ctx.readOnly(),
				structureLocked: workbook.structureLocked === true,
				visibleSheets: visible,
				hiddenSheets: workbook.sheets.length - visible,
				protected: workbook.sheets[sheet]?.protection?.sheet === true,
			},
			{ rename: () => rename(sheet) },
		);
		openContextMenu(ctx, entries, x, y, { restoreFocus: () => tabFor(strip, sheet)?.focus() });
	};

	strip.addEventListener('click', (event) => {
		const sheet = sheetOf(event.target);
		if (sheet !== undefined && !renaming) activate(sheet);
	});
	strip.addEventListener('dblclick', (event) => {
		const sheet = sheetOf(event.target);
		if (sheet !== undefined && !renaming) rename(sheet);
	});
	strip.addEventListener('contextmenu', (event) => {
		const sheet = sheetOf(event.target);
		if (sheet === undefined || renaming) return;
		event.preventDefault();
		openMenu(sheet, event.clientX, event.clientY);
	});
	strip.addEventListener('keydown', (event) => {
		const sheet = sheetOf(event.target);
		const workbook = ctx.workbook();
		if (sheet === undefined || !workbook || renaming) return;
		const visible = visibleSheetIndices(workbook);
		const at = visible.indexOf(sheet);
		const focusAt = (i: number) => {
			const target = visible[Math.max(0, Math.min(visible.length - 1, i))];
			if (target === undefined) return;
			focusable = target;
			for (const tab of strip.querySelectorAll<HTMLElement>('.xst-tab'))
				tab.tabIndex = tab.dataset.sheet === String(target) ? 0 : -1;
			scrollIntoView(target);
			tabFor(strip, target)?.focus();
		};
		const moves: Record<string, number> = {
			ArrowLeft: at - 1,
			ArrowRight: at + 1,
			Home: 0,
			End: visible.length - 1,
		};
		const target = moves[event.key];
		const menuKey = event.key === 'ContextMenu' || (event.key === 'F10' && event.shiftKey);
		if (target === undefined && !menuKey && !['Enter', ' ', 'F2'].includes(event.key)) return;
		event.preventDefault();
		event.stopPropagation();
		if (target !== undefined) focusAt(target);
		else if (event.key === 'F2') rename(sheet);
		else if (menuKey) {
			const box = (event.target as HTMLElement).getBoundingClientRect();
			openMenu(sheet, box.left, box.top);
		} else {
			activate(sheet);
			tabFor(strip, sheet)?.focus();
		}
	});
	strip.addEventListener(
		'wheel',
		(event) => {
			const delta = Math.abs(event.deltaX) > Math.abs(event.deltaY) ? event.deltaX : event.deltaY;
			if (!delta) return;
			event.preventDefault();
			strip.scrollLeft += delta;
		},
		{ passive: false },
	);
	const scrollBy = (dir: 1 | -1) => {
		const tabs = [...strip.querySelectorAll<HTMLElement>('.xst-tab')];
		const left = strip.scrollLeft;
		if (dir > 0) {
			const tab = tabs.find((t) => t.offsetLeft > left + 1);
			strip.scrollLeft = tab ? tab.offsetLeft : left;
		} else {
			const tab = [...tabs].reverse().find((t) => t.offsetLeft < left - 1);
			strip.scrollLeft = tab ? tab.offsetLeft : 0;
		}
	};
	prev.addEventListener('click', () => scrollBy(-1));
	next.addEventListener('click', () => scrollBy(1));
	add.addEventListener('click', () => {
		const session = ctx.session();
		if (!session || !editable()) return;
		try {
			const name = nextSheetName(session.workbook, sheetBaseName(ctx.t));
			const index = session.addSheet(name, ctx.activeSheet() + 1);
			activate(index);
		} catch (error) {
			ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'warning');
		}
	});

	const disposeDrag = wireTabDrag({
		strip,
		enabled: () => editable() && !renaming,
		visible: () => {
			const workbook = ctx.workbook();
			return workbook ? visibleSheetIndices(workbook) : [];
		},
		move: (from, to) => {
			try {
				ctx.session()?.moveSheet(from, to);
				activate(to);
			} catch (error) {
				ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'warning');
			}
		},
	});
	const disposeModel = ctx.onModelChange(() => render());
	render();

	return () => {
		renaming?.cancel();
		disposeModel();
		disposeDrag();
		root.remove();
	};
}
