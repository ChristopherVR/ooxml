// Hover tooltips for comments (the red corner triangle) and hyperlinks, and hyperlink following:
// external targets open only when `isSafeHyperlinkHref` accepts them; in-workbook locations
// (`Sheet2!B4`, defined names) navigate inside the editor.
import { rangeContains, sheetByName, type CellAddress } from '@christophervr/xlsx-core';
import { isSafeHyperlinkHref } from 'ooxml-core/opc';
import { h, viewOf } from './dom.js';
import { referenceTarget } from './formula-text.js';
import type { GridSelection } from './grid-selection.js';
import type { GridView } from './grid-view.js';
import { selectCell } from './selection-ops.js';

const HOVER_MS = 350;

export function wireTooltips(view: GridView): () => void {
	const tip = h(view.doc, 'div', 'xg-popup xg-tip', { role: 'tooltip' });
	tip.hidden = true;
	view.view.append(tip);
	let timer: ReturnType<typeof setTimeout> | undefined;
	let shownFor = '';
	const hide = () => {
		if (timer) clearTimeout(timer);
		timer = undefined;
		tip.hidden = true;
		shownFor = '';
	};
	const show = (at: CellAddress) => {
		const sheet = view.sheet();
		if (!sheet) return;
		const comment = sheet.comments.find(
			(c) => c.address.row === at.row && c.address.col === at.col,
		);
		const link = sheet.hyperlinks.find((l) => rangeContains(l.range, at));
		if (!comment && !link) return hide();
		tip.replaceChildren();
		if (comment) {
			const author = h(view.doc, 'b');
			author.textContent = `${comment.author}:`;
			tip.append(author, view.doc.createTextNode(comment.text));
			for (const reply of comment.replies ?? []) {
				const who = h(view.doc, 'b');
				who.textContent = `${reply.author}:`;
				tip.append(who, view.doc.createTextNode(reply.text));
			}
		} else if (link) {
			const target = link.target ?? link.location ?? '';
			tip.textContent = `${link.tooltip ?? target}\n${view.ctx.t('Ctrl+click to follow the link.')}`;
		}
		const g = view.geometry;
		const box = g.rangeBox({ start: at, end: at });
		tip.style.left = `${Math.min(box.x + box.w + 6, g.width - 200)}px`;
		tip.style.top = `${box.y}px`;
		tip.hidden = false;
	};
	const onMove = (event: PointerEvent) => {
		if (event.buttons) return hide();
		const hit = view.hitClient(event.clientX, event.clientY);
		if (hit.area !== 'cell') return hide();
		const key = `${hit.row}:${hit.col}`;
		if (key === shownFor) return;
		hide();
		const sheet = view.sheet();
		const has =
			sheet?.comments.some((c) => c.address.row === hit.row && c.address.col === hit.col) ||
			sheet?.hyperlinks.some((l) => rangeContains(l.range, hit));
		if (!has) return;
		shownFor = key;
		timer = setTimeout(() => show({ row: hit.row, col: hit.col }), HOVER_MS);
	};
	view.view.addEventListener('pointermove', onMove);
	view.view.addEventListener('pointerleave', hide);
	view.view.addEventListener('pointerdown', hide);
	view.scroller.addEventListener('scroll', hide, { passive: true });
	return () => {
		hide();
		view.view.removeEventListener('pointermove', onMove);
		view.view.removeEventListener('pointerleave', hide);
		view.view.removeEventListener('pointerdown', hide);
		view.scroller.removeEventListener('scroll', hide);
		tip.remove();
	};
}

/** Follows the hyperlink of a cell; false when the cell has none. */
export function followLink(
	view: GridView,
	selection: GridSelection,
	row: number,
	col: number,
): boolean {
	const sheet = view.sheet();
	const workbook = view.workbook();
	const link = sheet?.hyperlinks.find((l) => rangeContains(l.range, { row, col }));
	if (!link || !workbook) return false;
	if (link.target) {
		if (!isSafeHyperlinkHref(link.target)) {
			view.ctx.toast(
				view.ctx.t('This link was not opened because its address is not allowed.'),
				'warning',
			);
			return true;
		}
		viewOf(view.root)?.open(link.target, '_blank', 'noopener,noreferrer');
		return true;
	}
	const location = (link.location ?? '').replace(/^#/, '');
	const name = workbook.definedNames.find((n) => n.name.toLowerCase() === location.toLowerCase());
	const target = referenceTarget(name ? name.formula : location);
	if (!target) return true;
	const sheetIndex = target.sheet
		? workbook.sheets.indexOf(sheetByName(workbook, target.sheet) as never)
		: view.sheetIndex();
	if (sheetIndex < 0) return true;
	if (sheetIndex !== view.sheetIndex()) view.ctx.setActiveSheet(sheetIndex);
	const next = selectCell(sheetIndex, workbook.sheets[sheetIndex], target.range.start);
	selection.set({ ...next, ranges: [target.range] }, target.range.start);
	return true;
}
