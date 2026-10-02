// Test support (imported by *.test.ts only): a real EditorContext over an edit session, with the
// shell's command and dialog registries, so grid, formula bar, tab and menu tests drive the same
// paths the shell does.
import {
	createEditSession,
	createWorkbook,
	type CellRange,
	type EditSession,
	type Workbook,
} from '@christophervr/xlsx-core';
import { createCommandRegistry } from '../commands.js';
import type { EditorContext, GridController, Selection } from '../context.js';
import { createDialogRegistry } from '../dialogs.js';

export interface TestContext extends EditorContext {
	readonly events: { type: string; detail: unknown }[];
	readonly toasts: string[];
	setReadOnly(value: boolean): void;
	/** Simulates the shell's model-change fan-out. */
	notify(change?: unknown): void;
}

export function interpolate(text: string, vars?: Record<string, string | number>): string {
	return vars ? text.replace(/\{(\w+)\}/g, (all, name: string) => String(vars[name] ?? all)) : text;
}

export function createTestContext(
	workbook: Workbook = createWorkbook(),
	doc: Document = document,
): TestContext {
	const host = doc.createElement('div');
	doc.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const session: EditSession = createEditSession(workbook);
	let readOnly = false;
	let grid: GridController | undefined;
	let selection: Selection = {
		sheet: workbook.activeSheet,
		active: { row: 0, col: 0 },
		anchor: { row: 0, col: 0 },
		ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
	};
	const selectionListeners = new Set<(s: Selection) => void>();
	const modelListeners = new Set<(c: unknown) => void>();
	const events: { type: string; detail: unknown }[] = [];
	const toasts: string[] = [];
	const notify = (change?: unknown) => {
		for (const listener of [...modelListeners]) listener(change);
	};
	session.onChange((change) => notify(change));
	const ctx: TestContext = {
		host,
		root,
		events,
		toasts,
		session: () => session,
		workbook: () => session.workbook,
		activeSheet: () => session.workbook.activeSheet ?? 0,
		setActiveSheet(index) {
			const wb = session.workbook;
			if (!wb?.sheets[index] || wb.activeSheet === index) return;
			wb.activeSheet = index;
			events.push({ type: 'sheet-change', detail: { index, name: wb.sheets[index]?.name } });
			notify({ kind: 'sheet' });
		},
		selection: {
			get: () => selection,
			set(next: Partial<Selection> & { ranges?: CellRange[] }) {
				const moved = next.active ?? next.ranges ?? next.sheet;
				selection = { ...selection, ...next };
				if (selection.drawing === undefined || (moved !== undefined && !('drawing' in next)))
					delete selection.drawing;
				for (const listener of [...selectionListeners]) listener(selection);
			},
			onChange(listener) {
				selectionListeners.add(listener);
				return () => selectionListeners.delete(listener);
			},
		},
		readOnly: () => readOnly,
		setReadOnly(value) {
			readOnly = value;
		},
		t: interpolate,
		commands: createCommandRegistry(() => ctx),
		dialogs: createDialogRegistry(() => ctx),
		grid: () => grid,
		attachGrid(next) {
			grid = next;
		},
		onModelChange(listener) {
			modelListeners.add(listener);
			return () => modelListeners.delete(listener);
		},
		notify,
		requestRender() {},
		toast(message) {
			toasts.push(message);
		},
		emit(type, detail) {
			events.push({ type, detail });
			return true;
		},
		authorName: () => 'Tester',
		locale: () => 'en',
	};
	return ctx;
}
