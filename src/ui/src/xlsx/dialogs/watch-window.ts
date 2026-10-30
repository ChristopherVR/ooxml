// The Watch Window: a floating, non-modal list of watched cells (Book, Sheet, Name, Cell, Value,
// Formula) that stays current while the workbook changes. Rows come from the core's
// `describeWatch`; the watch list itself lives in the editor's audit state.
import { describeWatch, formatRange, watchExists, watchValueText } from 'ooxml-core/xlsx';
import {
	addWatch,
	auditState,
	onAuditChange,
	removeWatch,
	setWatchWindowOpen,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';
import { ensureStyle } from '../grid/dom';
import { el, textInput, field } from './fields';
import { button, showDialog } from './frame';
import { resolveReference, selectOn } from './go-to';

const MAX_WATCHES_AT_ONCE = 1000;

export const WATCH_CSS = `
.xve-watch{position:fixed;right:16px;bottom:40px;width:min(640px,92vw);height:240px;min-width:280px;min-height:140px;z-index:40;
	display:flex;flex-direction:column;resize:both;overflow:hidden;box-sizing:border-box;
	background:var(--xve-background,var(--office-background,#fff));color:var(--xve-foreground,var(--office-foreground,#201f1e));
	border:1px solid var(--xve-border,var(--office-border,#c8c6c4));border-radius:var(--xve-radius,var(--office-radius,4px));
	box-shadow:0 6px 24px rgba(0,0,0,.25);font:12px/1.4 var(--xve-font,var(--office-font,system-ui,sans-serif))}
.xve-watch[hidden]{display:none}
.xve-watch-head{display:flex;align-items:center;justify-content:space-between;padding:4px 8px;font-weight:600;
	border-bottom:1px solid var(--xve-border,var(--office-border,#c8c6c4))}
.xve-watch-bar{display:flex;gap:6px;padding:4px 8px}
.xve-watch-scroll{flex:1;overflow:auto}
.xve-watch table{border-collapse:collapse;width:100%}
.xve-watch th,.xve-watch td{text-align:left;padding:2px 8px;border-bottom:1px solid var(--xve-border,var(--office-border,#e1dfdd));white-space:nowrap;
	max-width:220px;overflow:hidden;text-overflow:ellipsis}
.xve-watch th{position:sticky;top:0;background:var(--xve-muted,var(--office-muted,#f3f2f1))}
.xve-watch tr[aria-selected=true]{background:var(--xve-accent,var(--office-selected,#e1ecf7))}
`;

const COLUMNS = ['Book', 'Sheet', 'Name', 'Cell', 'Value', 'Formula'] as const;

function openAddWatch(ctx: EditorContext): void {
	const selection = ctx.selection.get();
	const range = selection.ranges[0] ?? { start: selection.active, end: selection.active };
	const input = textInput(
		ctx,
		`${ctx.workbook()?.sheets[ctx.activeSheet()]?.name ?? ''}!${formatRange(range)}`,
	);
	void showDialog<true>(ctx, {
		name: 'add-watch',
		heading: 'Add Watch',
		okLabel: 'Add',
		body: [field(ctx, 'Choose the cells that you would like to watch the value of:', input)],
		opened: () => input.select(),
		submit: () => {
			const target = resolveReference(ctx, input.value);
			if (!target) {
				input.focus();
				return undefined;
			}
			let added = 0;
			for (let row = target.range.start.row; row <= target.range.end.row; row++)
				for (let col = target.range.start.col; col <= target.range.end.col; col++) {
					if (added++ >= MAX_WATCHES_AT_ONCE) return true;
					addWatch(ctx, { sheet: target.sheet, row, col });
				}
			return true;
		},
	});
}

/** Creates the window (hidden) in the editor's dialog area; returns a disposer. */
export function mountWatchWindow(ctx: EditorContext): () => void {
	ensureStyle(ctx.root, 'xve-watch', WATCH_CSS);
	const panel = el(ctx, 'div', 'xve-watch');
	panel.setAttribute('role', 'dialog');
	panel.dataset.part = 'watch-window';
	panel.hidden = true;
	const head = el(ctx, 'div', 'xve-watch-head');
	const title = el(ctx, 'span');
	title.textContent = ctx.t('Watch Window');
	const closeBtn = button(ctx, 'Close');
	head.append(title, closeBtn);
	const bar = el(ctx, 'div', 'xve-watch-bar');
	const add = button(ctx, 'Add Watch...');
	const del = button(ctx, 'Delete Watch');
	bar.append(add, del);
	const scroll = el(ctx, 'div', 'xve-watch-scroll');
	panel.append(head, bar, scroll);
	ctx.root.append(panel);
	let selected = -1;

	const render = (): void => {
		const state = auditState(ctx);
		panel.hidden = !state.watchOpen;
		if (!state.watchOpen) return;
		title.textContent = ctx.t('Watch Window');
		const workbook = ctx.workbook();
		const doc = ctx.host.ownerDocument;
		const table = doc.createElement('table');
		const headRow = table.createTHead().insertRow();
		for (const name of COLUMNS) {
			const th = doc.createElement('th');
			th.textContent = ctx.t(name);
			headRow.append(th);
		}
		const body = table.createTBody();
		const book = (ctx.host as { fileName?: string }).fileName || 'Book1';
		state.watches.forEach((watch, index) => {
			if (!workbook || !watchExists(workbook, watch)) return;
			const info = describeWatch(workbook, watch);
			const row = body.insertRow();
			row.setAttribute('aria-selected', String(index === selected));
			for (const value of [
				book,
				info.sheet,
				info.name,
				info.cell,
				watchValueText(info.value),
				info.formula,
			]) {
				const td = row.insertCell();
				td.textContent = value;
				td.title = value;
			}
			row.addEventListener('click', () => {
				selected = index;
				render();
			});
			row.addEventListener('dblclick', () =>
				selectOn(ctx, watch.sheet, { start: watch, end: watch }),
			);
		});
		scroll.replaceChildren(table);
		del.disabled = !state.watches[selected];
	};

	add.addEventListener('click', () => openAddWatch(ctx));
	del.addEventListener('click', () => {
		const watch = auditState(ctx).watches[selected];
		if (watch) removeWatch(ctx, watch);
		selected = -1;
	});
	closeBtn.addEventListener('click', () => setWatchWindowOpen(ctx, false));
	const offAudit = onAuditChange(ctx, render);
	const offModel = ctx.onModelChange(render);
	render();
	return () => {
		offAudit();
		offModel();
		panel.remove();
	};
}
