// Hover tooltips for comments (the red corner triangle) and hyperlinks, and hyperlink following:
// external targets open (in a new browsing context without opener or referrer) only when the
// core's `isOpenableHyperlinkHref` accepts them (http, https, mailto); in-workbook locations
// (`Sheet2!B4`, `#Sheet2!B4`, defined names), also when stored as a scheme-less target, navigate
// inside the editor.
import { rangeContains, sheetByName, type CellAddress } from '@christophervr/xlsx-core';
import { hyperlinkPolicy } from 'ooxml-core/opc';
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

/** The in-workbook location a scheme-less target names (`#Sheet2!A1`, `Sheet2!A1`), if any. */
function workbookLocation(target: string): string | undefined {
	const value = target.trim();
	if (value.startsWith('#')) return value.slice(1);
	return /^[a-z][a-z0-9+.-]+:/i.test(value) || /[\\/]/.test(value) ? undefined : value;
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
	const inside = link.target ? workbookLocation(link.target) : undefined;
	if (link.target && !(inside && goTo(view, selection, inside))) {
		const policy = hyperlinkPolicy(link.target);
		if (policy.open) viewOf(view.root)?.open(link.target, '_blank', 'noopener,noreferrer');
		else
			view.ctx.toast(
				view.ctx.t(
					policy.store
						? 'This link is kept in the workbook, but only web and e-mail links open from the editor.'
						: 'This link was not opened because its address is not allowed.',
				),
				'warning',
			);
		return true;
	}
	if (!link.target) goTo(view, selection, link.location ?? '');
	return true;
}

/** Selects an in-workbook location (`Sheet2!B4`, `'My sheet'!A1`, a defined name); false when it names none. */
function goTo(view: GridView, selection: GridSelection, location: string): boolean {
	const workbook = view.workbook();
	if (!workbook) return false;
	const text = location.replace(/^#/, '');
	const name = workbook.definedNames.find((n) => n.name.toLowerCase() === text.toLowerCase());
	const target = referenceTarget(name ? name.formula : text);
	if (!target) return false;
	const sheetIndex = target.sheet
		? workbook.sheets.indexOf(sheetByName(workbook, target.sheet) as never)
		: view.sheetIndex();
	if (sheetIndex < 0) return false;
	if (sheetIndex !== view.sheetIndex()) view.ctx.setActiveSheet(sheetIndex);
	const next = selectCell(sheetIndex, workbook.sheets[sheetIndex], target.range.start);
	selection.set({ ...next, ranges: [target.range] }, target.range.start);
	return true;
}
