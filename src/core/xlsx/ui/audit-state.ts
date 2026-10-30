// Formula auditing state of one editor: the trace arrows drawn over the grid and the Watch Window's
// cells. Held per editor context (not in the workbook, as in Excel the arrows are a view overlay).
import { DEFAULT_ERROR_RULES, type CellRef, type ErrorRule, type TraceArrow } from '../audit/index';
import type { EditorContext } from './context';

export interface AuditState {
	arrows: TraceArrow[];
	watches: CellRef[];
	watchOpen: boolean;
	/** Error Checking rules switched on (File > Options > Formulas). */
	rules: ErrorRule[];
	listeners: Set<() => void>;
}

const states = new WeakMap<EditorContext, AuditState>();

export function auditState(ctx: EditorContext): AuditState {
	let state = states.get(ctx);
	if (!state) {
		const created: AuditState = {
			arrows: [],
			watches: [],
			watchOpen: false,
			rules: [...DEFAULT_ERROR_RULES],
			listeners: new Set(),
		};
		states.set(ctx, (state = created));
		// Excel drops the arrows when rows, columns, sheets or the cells' order change.
		ctx.onModelChange((change) => {
			const c = change as { kind?: string; label?: string; structural?: boolean } | undefined;
			if (!created.arrows.length || !c) return;
			if (
				c.structural ||
				c.kind === 'structure' ||
				c.kind === 'sheets' ||
				/^sort/i.test(c.label ?? '')
			) {
				created.arrows = [];
				notify(created);
				ctx.grid()?.invalidate();
				ctx.requestRender();
			}
		});
	}
	return state;
}

function notify(state: AuditState): void {
	for (const listener of state.listeners) listener();
}

/** Calls `listener` when arrows, watches or the Watch Window's visibility change. */
export function onAuditChange(ctx: EditorContext, listener: () => void): () => void {
	const state = auditState(ctx);
	state.listeners.add(listener);
	return () => state.listeners.delete(listener);
}

export function setTraceArrows(ctx: EditorContext, arrows: TraceArrow[]): void {
	const state = auditState(ctx);
	state.arrows = arrows;
	notify(state);
	ctx.grid()?.invalidate();
	ctx.requestRender();
}

export function setErrorRules(ctx: EditorContext, rules: ErrorRule[]): void {
	const state = auditState(ctx);
	state.rules = rules;
	notify(state);
}

export function addWatch(ctx: EditorContext, cell: CellRef): void {
	const state = auditState(ctx);
	if (state.watches.some((w) => w.sheet === cell.sheet && w.row === cell.row && w.col === cell.col))
		return;
	state.watches = [...state.watches, cell];
	notify(state);
}

export function removeWatch(ctx: EditorContext, cell: CellRef): void {
	const state = auditState(ctx);
	state.watches = state.watches.filter(
		(w) => !(w.sheet === cell.sheet && w.row === cell.row && w.col === cell.col),
	);
	notify(state);
}

export function setWatchWindowOpen(ctx: EditorContext, open: boolean): void {
	const state = auditState(ctx);
	state.watchOpen = open;
	notify(state);
	ctx.requestRender();
}
