// Test double of the shell's EditorContext over a real EditSession (used by command and dialog
// tests only). Translation is the identity with `{name}` placeholders filled.
import {
	type CellAddress,
	type CellRange,
	type EditSession,
	type Workbook,
	createEditSession,
	createWorkbook,
	parseRange,
} from '@christophervr/xlsx-core';
import { createCommandRegistry } from '../commands.js';
import type { EditorContext, GridController, Selection } from '../context.js';
import { createDialogRegistry } from '../dialogs.js';

export interface TestContext extends EditorContext {
	select(ref: string): void;
	toasts: { message: string; kind?: string }[];
	events: { type: string; detail: unknown }[];
	setReadOnly(value: boolean): void;
}

export function createTestContext(
	workbook: Workbook = createWorkbook(),
	options: { grid?: Partial<GridController>; dom?: boolean } = {},
): TestContext {
	let session: EditSession | undefined = createEditSession(workbook);
	let readOnly = false;
	const doc = typeof document === 'undefined' ? undefined : document;
	const host = (doc ? doc.createElement('div') : ({} as HTMLElement)) as HTMLElement;
	if (doc) doc.body.append(host);
	const root = (doc ? host.attachShadow({ mode: 'open' }) : ({} as ShadowRoot)) as ShadowRoot;
	const origin: CellAddress = { row: 0, col: 0 };
	let selection: Selection = {
		sheet: 0,
		active: origin,
		anchor: origin,
		ranges: [{ start: origin, end: origin }],
	};
	const selectionListeners = new Set<(s: Selection) => void>();
	const modelListeners = new Set<(c: unknown) => void>();
	let grid: GridController | undefined = options.grid
		? ({
				focus() {},
				scrollTo() {},
				invalidate() {},
				beginEdit() {},
				commitEdit: () => true,
				cancelEdit() {},
				isEditing: () => false,
				measureText: (text: string) => text.length * 7,
				zoom: () => 100,
				setZoom() {},
				...options.grid,
			} as GridController)
		: undefined;
	session.onChange((change) => {
		for (const l of modelListeners) l(change);
	});
	const ctx: TestContext = {
		host,
		root,
		session: () => session,
		workbook: () => session?.workbook,
		activeSheet: () => selection.sheet,
		setActiveSheet(index) {
			selection = { ...selection, sheet: index };
			for (const l of modelListeners) l({ kind: 'sheet' });
		},
		selection: {
			get: () => selection,
			set(next) {
				const moved = next.active ?? next.ranges ?? next.sheet;
				selection = { ...selection, ...next };
				if (selection.drawing === undefined || (moved !== undefined && !('drawing' in next)))
					delete selection.drawing;
				for (const l of selectionListeners) l(selection);
			},
			onChange(listener) {
				selectionListeners.add(listener);
				return () => selectionListeners.delete(listener);
			},
		},
		readOnly: () => readOnly,
		t: (key, vars) => key.replace(/\{(\w+)\}/g, (m, name: string) => String(vars?.[name] ?? m)),
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
		requestRender() {},
		toast(message, kind) {
			ctx.toasts.push(kind ? { message, kind } : { message });
		},
		emit(type, detail) {
			ctx.events.push({ type, detail });
			return true;
		},
		authorName: () => 'Tester',
		locale: () => 'en',
		select(ref) {
			const range: CellRange | undefined = parseRange(ref);
			if (!range) throw new Error(`Bad range ${ref}`);
			ctx.selection.set({ active: range.start, anchor: range.start, ranges: [range] });
		},
		toasts: [],
		events: [],
		setReadOnly(value) {
			readOnly = value;
		},
	};
	if (!workbook) session = undefined;
	return ctx;
}

/** The open dialog element by `data-dialog` name. */
export function dialogEl(ctx: EditorContext, name: string): HTMLElement {
	const el = ctx.root.querySelector<HTMLElement>(`[data-dialog="${name}"]`);
	if (!el) throw new Error(`Dialog ${name} is not open`);
	return el;
}

export function clickButton(root: HTMLElement, label: string): void {
	const b = [...root.querySelectorAll('button')].find((x) => x.textContent === label);
	if (!b) throw new Error(`No button ${label}`);
	b.click();
}

export function inputByLabel<T extends HTMLElement = HTMLInputElement>(
	root: HTMLElement,
	label: string,
): T {
	const el = root.querySelector<T>(`[aria-label="${label}"]`);
	if (!el) throw new Error(`No field ${label}`);
	return el;
}

export function setValue(
	el: HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement,
	value: string,
): void {
	el.value = value;
	el.dispatchEvent(new Event('input', { bubbles: true }));
	el.dispatchEvent(new Event('change', { bubbles: true }));
}

export function pressKey(el: HTMLElement, key: string): void {
	el.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
}

export const tick = (): Promise<void> => new Promise((r) => setTimeout(r, 0));
