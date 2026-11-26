import { formatAddress, type CellAddress, type Comment, type Worksheet } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { h } from './dom';
import type { GridView } from './grid-view';

const showAll = new WeakMap<EditorContext, WeakSet<Worksheet>>();
const navigated = new WeakMap<EditorContext, { sheet: Worksheet; at: CellAddress }>();
const activeSheet = (ctx: EditorContext) => ctx.workbook()?.sheets[ctx.activeSheet()];

export const isNote = (comment: Comment): boolean => comment.replies === undefined;

export function allNotesShown(ctx: EditorContext): boolean {
	const sheet = activeSheet(ctx);
	return sheet !== undefined && (showAll.get(ctx)?.has(sheet) ?? false);
}

export function toggleAllNotes(ctx: EditorContext): void {
	const sheet = activeSheet(ctx);
	if (!sheet) return;
	let sheets = showAll.get(ctx);
	if (!sheets) showAll.set(ctx, (sheets = new WeakSet()));
	if (sheets.has(sheet)) sheets.delete(sheet);
	else sheets.add(sheet);
	navigated.delete(ctx);
	ctx.grid()?.invalidate();
	ctx.requestRender();
}

export function showNavigatedNote(ctx: EditorContext, sheet: Worksheet, at: CellAddress): void {
	navigated.set(ctx, { sheet, at });
	ctx.grid()?.invalidate();
}

export function renderAnnotation(popup: HTMLElement, comment: Comment): void {
	const { ownerDocument: doc } = popup;
	popup.replaceChildren();
	for (const record of [comment, ...(comment.replies ?? [])]) {
		const author = h(doc, 'b');
		author.textContent = `${record.author}:`;
		popup.append(author, doc.createTextNode(record.text));
	}
}

export function placeAnnotationPopup(
	popup: HTMLElement,
	box: { x: number; y: number; w: number },
	width: number,
	height: number,
): void {
	popup.style.left = `${Math.max(0, Math.min(box.x + box.w + 6, width - popup.offsetWidth))}px`;
	popup.style.top = `${Math.max(0, Math.min(box.y, height - popup.offsetHeight))}px`;
}

export function installNotePopups(view: GridView): () => void {
	const { ctx } = view;
	const popups = new Map<string, HTMLElement>();
	const paint = () => {
		const sheet = view.sheet();
		const g = view.geometry;
		const selected = navigated.get(ctx);
		const all = allNotesShown(ctx);
		if (!all && selected?.sheet !== sheet && !popups.size) return;
		const showing = new Set<string>();
		if (sheet && (all || selected?.sheet === sheet)) {
			const { rows, cols, frozenRows, frozenCols } = view.visible;
			const visibleRows = new Set([...rows, ...frozenRows]);
			const visibleCols = new Set([...cols, ...frozenCols]);
			for (const note of sheet.comments) {
				const at = note.address;
				if (
					!isNote(note) ||
					!visibleRows.has(at.row) ||
					!visibleCols.has(at.col) ||
					(!all && (at.row !== selected?.at.row || at.col !== selected.at.col))
				)
					continue;
				const box = g.rangeBox({ start: at, end: at });
				if (box.w <= 0 || box.h <= 0 || box.x >= g.width || box.y >= g.height) continue;
				const key = `${at.row}:${at.col}`;
				showing.add(key);
				let popup = popups.get(key);
				if (!popup) {
					popup = h(view.doc, 'div', 'xg-popup xg-tip', {
						role: 'note',
						'aria-label': formatAddress(at),
						'data-note': key,
					});
					popups.set(key, popup);
					view.view.append(popup);
				}
				renderAnnotation(popup, note);
				placeAnnotationPopup(popup, box, g.width, g.height);
			}
		}
		for (const [key, popup] of popups) {
			if (showing.has(key)) continue;
			popup.remove();
			popups.delete(key);
		}
	};
	view.hooks.add(paint);
	const offSelection = ctx.selection.onChange(({ active, sheet }) => {
		const note = navigated.get(ctx);
		if (
			note &&
			(note.sheet !== ctx.workbook()?.sheets[sheet] ||
				note.at.row !== active.row ||
				note.at.col !== active.col)
		)
			navigated.delete(ctx);
	});
	return () => {
		offSelection();
		view.hooks.delete(paint);
		for (const popup of popups.values()) popup.remove();
	};
}
