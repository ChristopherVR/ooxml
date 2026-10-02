// Formulas tab: Insert Function, the Function Library category menus (from the core
// FUNCTION_CATALOG), Defined Names, Show Formulas and Calculation.
import { FUNCTION_CATALOG, type FunctionInfo } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { icon } from './icons.js';
import { editing, refresh, target, viewing } from './util.js';

/** Function Library menus: [menu label, catalog category]. */
export const FUNCTION_MENUS: ReadonlyArray<
	readonly [label: string, category: string, glyph: string]
> = [
	['Financial', 'Financial', 'financial'],
	['Logical', 'Logical', 'logical'],
	['Text', 'Text', 'text'],
	['Date & Time', 'Date & Time', 'dateTime'],
	['Lookup & Reference', 'Lookup & Reference', 'lookup'],
	['Math & Trig', 'Math & Trig', 'math'],
];

/** Categories under More Functions. */
export const MORE_FUNCTION_CATEGORIES = ['Statistical', 'Engineering', 'Information'] as const;

export const functionsIn = (category: string): readonly FunctionInfo[] =>
	FUNCTION_CATALOG.filter((f) => f.category === category);

const recent: string[] = [
	'SUM',
	'AVERAGE',
	'IF',
	'HYPERLINK',
	'COUNT',
	'MAX',
	'SIN',
	'SUMIF',
	'PMT',
	'STDEV',
];

/** Most recently used functions, newest first (Excel keeps ten). */
export const recentFunctions = (): readonly string[] =>
	recent.filter((name) => FUNCTION_CATALOG.some((f) => f.name === name));

export function noteFunctionUse(name: string): void {
	const i = recent.indexOf(name);
	if (i >= 0) recent.splice(i, 1);
	recent.unshift(name);
	recent.length = Math.min(recent.length, 10);
}

/** Starts editing the active cell with `=NAME(` (or writes `=NAME()` when no grid is mounted). */
export function insertFunctionCall(ctx: EditorContext, name: string, args?: string): void {
	const t = target(ctx);
	if (!t) return;
	noteFunctionUse(name);
	const grid = ctx.grid();
	if (args !== undefined || !grid) {
		t.session.setCellInput(t.sheet, t.active.row, t.active.col, `=${name}(${args ?? ''})`);
		return;
	}
	if (grid.isEditing()) grid.commitEdit();
	grid.beginEdit(`=${name}(`);
}

export function formulaCommands(): Command[] {
	return [
		editing({
			id: 'formulas.insert-function',
			label: 'Insert Function',
			icon: icon('function'),
			shortcut: 'Shift+F3',
			lock: false,
			run: (ctx, arg) => void ctx.dialogs.open('insert-function', arg),
		}),
		editing({
			id: 'formulas.recent',
			label: 'Recently Used',
			icon: icon('recent'),
			lock: false,
			run: (ctx) => void ctx.dialogs.open('insert-function', { category: 'recent' }),
		}),
		editing({
			id: 'formulas.function',
			label: 'Insert Function',
			icon: icon('function'),
			lock: false,
			run: (ctx, arg) => {
				if (typeof arg === 'string' && FUNCTION_CATALOG.some((f) => f.name === arg))
					insertFunctionCall(ctx, arg);
			},
		}),
		viewing({
			id: 'formulas.name-manager',
			label: 'Name Manager',
			icon: icon('nameManager'),
			shortcut: 'Ctrl+F3',
			run: (ctx) => void ctx.dialogs.open('name-manager'),
		}),
		editing({
			id: 'formulas.define-name',
			label: 'Define Name',
			icon: icon('defineName'),
			lock: false,
			run: (ctx) => void ctx.dialogs.open('define-name'),
		}),
		editing({
			id: 'formulas.use-in-formula',
			label: 'Use in Formula',
			icon: icon('useInFormula'),
			lock: false,
			enabled: (ctx) =>
				!!ctx.workbook()?.definedNames.some((n) => !n.hidden && !n.name.startsWith('_xlnm.')),
			run: (ctx) => void ctx.dialogs.open('use-in-formula'),
		}),
		editing({
			id: 'formulas.create-from-selection',
			label: 'Create from Selection',
			icon: icon('fromSelection'),
			shortcut: 'Ctrl+Shift+F3',
			lock: false,
			run: (ctx) => void ctx.dialogs.open('create-names'),
		}),
		viewing({
			id: 'formulas.show-formulas',
			label: 'Show Formulas',
			icon: icon('showFormulas'),
			shortcut: 'Ctrl+`',
			checked: (ctx) => !!target(ctx)?.ws.view.showFormulas,
			run: (ctx) => {
				const t = target(ctx);
				if (t) t.session.setSheetView(t.sheet, { showFormulas: !t.ws.view.showFormulas });
			},
		}),
		...(['auto', 'manual'] as const).map((mode) =>
			editing({
				id: mode === 'auto' ? 'formulas.calc-automatic' : 'formulas.calc-manual',
				label: mode === 'auto' ? 'Automatic' : 'Manual',
				icon: icon('calculator'),
				lock: false,
				checked: (ctx) => (ctx.workbook()?.calcMode === 'manual' ? 'manual' : 'auto') === mode,
				run: (ctx) => ctx.session()?.setCalcMode(mode),
			}),
		),
		viewing({
			id: 'formulas.calculate-now',
			label: 'Calculate Now',
			icon: icon('calcNow'),
			shortcut: 'F9',
			run: (ctx) => {
				ctx.session()?.calculateNow();
				refresh(ctx);
			},
		}),
		viewing({
			id: 'formulas.calculate-sheet',
			label: 'Calculate Sheet',
			icon: icon('calcNow'),
			shortcut: 'Shift+F9',
			run: (ctx) => {
				ctx.session()?.calculateSheet(ctx.activeSheet());
				refresh(ctx);
			},
		}),
	];
}
