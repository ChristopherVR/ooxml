// @vitest-environment jsdom
import { createWorkbook, getCell, putCell, type Workbook } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import { editBridge } from './edit-bridge.js';
import { mountGrid } from './index.js';
import { createTestContext, type TestContext } from './test-context.js';

let dispose: (() => void) | undefined;
afterEach(() => {
	dispose?.();
	dispose = undefined;
	document.body.replaceChildren();
});

function setup(workbook: Workbook = createWorkbook()) {
	const ctx = createTestContext(workbook);
	const container = document.createElement('div');
	ctx.root.append(container);
	dispose = mountGrid(ctx, container);
	const sink = ctx.root.querySelector<HTMLTextAreaElement>('.xg-cell-field textarea')!;
	const key = (k: string, init: KeyboardEventInit = {}) =>
		sink.dispatchEvent(
			new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init }),
		);
	const type = (text: string) => {
		sink.value = text;
		sink.setSelectionRange(text.length, text.length);
		sink.dispatchEvent(new Event('input', { bubbles: true }));
	};
	return { ctx, sink, key, type };
}

const active = (ctx: TestContext) => ctx.selection.get().active;
const frame = () => new Promise((resolve) => setTimeout(resolve, 40));

describe('mountGrid', () => {
	it('renders headers, cells and the grid role, and attaches a controller', () => {
		const wb = createWorkbook();
		putCell(wb.sheets[0]!, 0, 0, { value: 'Hello' });
		const { ctx } = setup(wb);
		const grid = ctx.root.querySelector('[role="grid"]')!;
		expect(grid.getAttribute('aria-rowcount')).toBe('1048576');
		expect(ctx.root.querySelector('.xg-c .xg-tx')?.textContent).toBe('Hello');
		expect(
			[...ctx.root.querySelectorAll('.xg-hdr-col .xg-hd')].some((n) => n.textContent === 'A'),
		).toBe(true);
		expect(ctx.grid()).toBeDefined();
		expect(ctx.commands.get('edit.copy')).toBeDefined();
	});

	it('moves with arrows and extends with Shift, leaving selection-change to the core', () => {
		const { ctx, key } = setup();
		key('ArrowDown');
		key('ArrowRight');
		expect(active(ctx)).toEqual({ row: 1, col: 1 });
		key('ArrowRight', { shiftKey: true });
		expect(ctx.selection.get().ranges[0]).toEqual({
			start: { row: 1, col: 1 },
			end: { row: 1, col: 2 },
		});
		// The editor core announces selection-change (once, one formatter); the grid never does.
		expect(ctx.events.filter((e) => e.type === 'selection-change')).toEqual([]);
	});

	it('jumps with Ctrl+arrows using the core navigation', () => {
		const wb = createWorkbook();
		for (let r = 0; r < 5; r++) putCell(wb.sheets[0]!, r, 0, { value: r });
		const { ctx, key } = setup(wb);
		key('ArrowDown', { ctrlKey: true });
		expect(active(ctx)).toEqual({ row: 4, col: 0 });
		key('Home', { ctrlKey: true });
		expect(active(ctx)).toEqual({ row: 0, col: 0 });
	});

	it('types into a cell, commits with Enter and moves down', () => {
		const { ctx, type, key } = setup();
		type('42');
		expect(ctx.grid()?.isEditing()).toBe(true);
		expect(editBridge(ctx).state().text).toBe('42');
		key('Enter');
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 0)?.value).toBe(42);
		expect(active(ctx)).toEqual({ row: 1, col: 0 });
		expect(ctx.grid()?.isEditing()).toBe(false);
	});

	it('recalculates formulas after editing a precedent', () => {
		const { ctx, type, key } = setup();
		type('2');
		key('Enter');
		type('=A1*3');
		key('Enter');
		const sheet = ctx.workbook()!.sheets[0]!;
		expect(getCell(sheet, 1, 0)?.value).toBe(6);
		ctx.selection.set({
			active: { row: 0, col: 0 },
			anchor: { row: 0, col: 0 },
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
		});
		type('5');
		key('Tab');
		expect(getCell(sheet, 1, 0)?.value).toBe(15);
		expect(active(ctx)).toEqual({ row: 0, col: 1 });
	});

	it('cancels with Escape and edits the existing value with F2', () => {
		const wb = createWorkbook();
		putCell(wb.sheets[0]!, 0, 0, { value: 'keep' });
		const { ctx, type, key } = setup(wb);
		type('x');
		key('Escape');
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 0)?.value).toBe('keep');
		key('F2');
		expect(editBridge(ctx).state()).toMatchObject({ editing: true, mode: 'edit', text: 'keep' });
	});

	it('inserts pointed references with arrows in a formula (point mode)', async () => {
		const { ctx, type, key } = setup();
		ctx.selection.set({
			active: { row: 2, col: 2 },
			anchor: { row: 2, col: 2 },
			ranges: [{ start: { row: 2, col: 2 }, end: { row: 2, col: 2 } }],
		});
		type('=');
		key('ArrowUp');
		expect(editBridge(ctx).state().text).toBe('=C2');
		expect(editBridge(ctx).state().mode).toBe('point');
		key('ArrowUp', { shiftKey: true });
		expect(editBridge(ctx).state().text).toBe('=C1:C2');
		await frame();
		expect(ctx.root.querySelectorAll('.xg-refbox:not([hidden])').length).toBeGreaterThan(0);
	});

	it('clears contents with Delete and fills down with Ctrl+D', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 7 });
		const { ctx, key } = setup(wb);
		key('ArrowDown', { shiftKey: true });
		key('ArrowDown', { shiftKey: true });
		key('d', { ctrlKey: true });
		expect(getCell(ctx.workbook()!.sheets[0]!, 2, 0)?.value).toBe(7);
		key('Delete');
		expect(getCell(ctx.workbook()!.sheets[0]!, 2, 0)?.value ?? null).toBeNull();
	});

	it('does not edit in read-only mode but still navigates', () => {
		const { ctx, type, key } = setup();
		ctx.setReadOnly(true);
		type('9');
		expect(ctx.grid()?.isEditing()).toBe(false);
		key('ArrowDown');
		expect(active(ctx).row).toBe(1);
	});

	it('rejects a value that fails a stop validation and shows the alert', async () => {
		const wb = createWorkbook();
		wb.sheets[0]!.dataValidations.push({
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '10',
			errorStyle: 'stop',
			showErrorMessage: true,
			error: 'Between 1 and 10',
		});
		const { ctx, type, key } = setup(wb);
		type('50');
		key('Enter');
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 0)).toBeUndefined();
		const alert = ctx.root.querySelector('[role="alertdialog"]')!;
		expect(alert.textContent).toContain('Between 1 and 10');
		alert.querySelector<HTMLButtonElement>('[data-answer="cancel"]')!.click();
		await Promise.resolve();
		expect(ctx.grid()?.isEditing()).toBe(false);
	});

	it('copies and pastes through native clipboard events with the internal payload', async () => {
		const wb = createWorkbook();
		putCell(wb.sheets[0]!, 0, 0, { value: 'copied' });
		const { ctx, sink, key } = setup(wb);
		key('c', { ctrlKey: true });
		const data = new Map<string, string>();
		const clipboardData = {
			setData: (t: string, v: string) => data.set(t, v),
			getData: (t: string) => data.get(t) ?? '',
		};
		const copy = new Event('copy', { bubbles: true, cancelable: true });
		Object.defineProperty(copy, 'clipboardData', { value: clipboardData });
		sink.dispatchEvent(copy);
		expect(data.get('text/plain')).toContain('copied');
		expect(data.get('text/html')).toContain('<table');
		await frame();
		expect(ctx.root.querySelector('.xg-ants:not([hidden])')).not.toBeNull();
		key('ArrowRight');
		const paste = new Event('paste', { bubbles: true, cancelable: true });
		Object.defineProperty(paste, 'clipboardData', { value: clipboardData });
		sink.dispatchEvent(paste);
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 1)?.value).toBe('copied');
		key('Escape');
		await frame();
		expect(ctx.root.querySelector('.xg-ants:not([hidden])')).toBeNull();
	});
});
