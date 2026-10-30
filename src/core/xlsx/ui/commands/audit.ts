// Formulas > Formula Auditing: Trace Precedents / Dependents, Remove Arrows, Evaluate Formula,
// Error Checking and Watch Window. The logic is in `audit/`; the dialogs and the arrow layer
// are the UI package's.
import { removeArrows, traceNext, type TraceKind } from '../../audit/index';
import { auditState, setTraceArrows, setWatchWindowOpen } from '../audit-state';
import type { Command } from '../commands';
import type { EditorContext } from '../context';
import { icon } from './icons';
import { target, viewing } from './util';

const anchorOf = (ctx: EditorContext) => {
	const t = target(ctx);
	return t
		? { sheet: t.sheet, row: t.active.row, col: t.active.col, workbook: t.workbook }
		: undefined;
};

const MESSAGES: Record<string, string> = {
	'no-formula':
		'The Trace Precedents command requires that the active cell contain a formula that refers to other cells.',
	'no-dependents':
		'The Trace Dependents command requires that the active cell be referred to by formulas in the workbook.',
	'no-more': 'There are no more cells to trace at this level.',
};

function trace(ctx: EditorContext, kind: TraceKind): void {
	const a = anchorOf(ctx);
	if (!a) return;
	const { workbook, ...cell } = a;
	const result = traceNext(workbook, auditState(ctx).arrows, kind, cell);
	if (result.outcome !== 'ok') {
		ctx.toast(ctx.t(MESSAGES[result.outcome] ?? ''), 'warning');
		return;
	}
	setTraceArrows(ctx, result.arrows);
}

function remove(ctx: EditorContext, mode: 'all' | TraceKind): void {
	const a = anchorOf(ctx);
	setTraceArrows(
		ctx,
		removeArrows(
			auditState(ctx).arrows,
			mode,
			a ? { sheet: a.sheet, row: a.row, col: a.col } : undefined,
		),
	);
}

export function auditCommands(): Command[] {
	return [
		viewing({
			id: 'formulas.trace-precedents',
			label: 'Trace Precedents',
			icon: icon('tracePrecedents'),
			run: (ctx) => trace(ctx, 'precedents'),
		}),
		viewing({
			id: 'formulas.trace-dependents',
			label: 'Trace Dependents',
			icon: icon('traceDependents'),
			run: (ctx) => trace(ctx, 'dependents'),
		}),
		viewing({
			id: 'formulas.remove-arrows',
			label: 'Remove Arrows',
			icon: icon('removeArrows'),
			enabled: (ctx) => auditState(ctx).arrows.length > 0,
			run: (ctx) => remove(ctx, 'all'),
		}),
		viewing({
			id: 'formulas.remove-precedent-arrows',
			label: 'Remove Precedent Arrows',
			icon: icon('removeArrows'),
			enabled: (ctx) => auditState(ctx).arrows.some((a) => a.kind === 'precedents'),
			run: (ctx) => remove(ctx, 'precedents'),
		}),
		viewing({
			id: 'formulas.remove-dependent-arrows',
			label: 'Remove Dependent Arrows',
			icon: icon('removeArrows'),
			enabled: (ctx) => auditState(ctx).arrows.some((a) => a.kind === 'dependents'),
			run: (ctx) => remove(ctx, 'dependents'),
		}),
		viewing({
			id: 'formulas.evaluate-formula',
			label: 'Evaluate Formula',
			icon: icon('evaluateFormula'),
			run: (ctx) => void ctx.dialogs.open('evaluate-formula'),
		}),
		viewing({
			id: 'formulas.error-checking',
			label: 'Error Checking',
			icon: icon('errorChecking'),
			run: (ctx) => void ctx.dialogs.open('error-checking'),
		}),
		viewing({
			id: 'formulas.watch-window',
			label: 'Watch Window',
			icon: icon('watchWindow'),
			checked: (ctx) => auditState(ctx).watchOpen,
			run: (ctx) => setWatchWindowOpen(ctx, !auditState(ctx).watchOpen),
		}),
	];
}
