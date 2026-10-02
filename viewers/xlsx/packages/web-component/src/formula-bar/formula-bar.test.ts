// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook, createWorksheet, putCell, type Workbook } from '@christophervr/xlsx-core';
import { createTestContext, type TestContext } from '../grid/test-context.js';
import { editBridge, type BeginOptions, type EditDriver } from '../grid/edit-bridge.js';
import { mountFormulaBar } from './index.js';
import { absoluteReference, resolveNameBox } from './name-resolve.js';

const disposers: (() => void)[] = [];
afterEach(() => {
	while (disposers.length) disposers.pop()?.();
	document.body.replaceChildren();
});

function twoSheets(): Workbook {
	const wb = createWorkbook();
	wb.sheets.push(createWorksheet('Data Sheet', 2));
	putCell(wb.sheets[0]!, 0, 0, { value: 5, formula: '2+3' });
	wb.definedNames.push({ name: 'Totals', formula: "'Data Sheet'!$B$2:$C$4" });
	return wb;
}

function mount(wb = twoSheets()) {
	const ctx = createTestContext(wb);
	const container = document.createElement('div');
	ctx.root.append(container);
	disposers.push(mountFormulaBar(ctx, container));
	const q = <T extends Element>(sel: string) => ctx.root.querySelector<T>(sel)!;
	return {
		ctx,
		q,
		nameInput: q<HTMLInputElement>('[part="name-box"]'),
		formula: q<HTMLTextAreaElement>('.xfb-input textarea'),
	};
}

const typeName = (input: HTMLInputElement, text: string) => {
	input.value = text;
	input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
};

function fakeDriver(ctx: TestContext) {
	const calls: string[] = [];
	const bridge = editBridge(ctx);
	const driver: EditDriver = {
		begin(options: BeginOptions) {
			calls.push(`begin:${options.mode}:${options.source}:${options.text ?? ''}`);
			bridge.set(
				{
					editing: true,
					mode: options.mode ?? 'edit',
					text: options.text ?? '',
					caret: options.caret ?? 0,
					source: options.source ?? 'cell',
				},
				'grid',
			);
			return true;
		},
		commit(move, all) {
			calls.push(`commit:${move}:${all}`);
			bridge.set({ editing: false, mode: 'ready' }, 'grid');
			return true;
		},
		cancel() {
			calls.push('cancel');
			bridge.set({ editing: false, mode: 'ready' }, 'grid');
		},
		changed(source) {
			calls.push(`changed:${source}:${bridge.state().text}`);
		},
	};
	disposers.push(bridge.attach(driver));
	return { calls, bridge };
}

describe('name box', () => {
	it('shows the active cell and follows the selection', () => {
		const { ctx, nameInput } = mount();
		expect(nameInput.value).toBe('A1');
		ctx.selection.set({ active: { row: 2, col: 1 } });
		expect(nameInput.value).toBe('B3');
	});

	it('goes to a cell or a range', () => {
		const { ctx, nameInput } = mount();
		typeName(nameInput, 'B2:C3');
		const sel = ctx.selection.get();
		expect(sel.active).toEqual({ row: 1, col: 1 });
		expect(sel.ranges).toEqual([{ start: { row: 1, col: 1 }, end: { row: 2, col: 2 } }]);
		typeName(nameInput, 'd7');
		expect(ctx.selection.get().active).toEqual({ row: 6, col: 3 });
	});

	it('goes to a defined name on another sheet', () => {
		const { ctx, nameInput } = mount();
		typeName(nameInput, 'totals');
		expect(ctx.activeSheet()).toBe(1);
		expect(ctx.selection.get()).toMatchObject({
			sheet: 1,
			ranges: [{ start: { row: 1, col: 1 }, end: { row: 3, col: 2 } }],
		});
		expect(nameInput.value).toBe('Totals');
	});

	it('goes to a sheet-qualified reference', () => {
		const { ctx, nameInput } = mount();
		typeName(nameInput, "'Data Sheet'!C5");
		expect(ctx.activeSheet()).toBe(1);
		expect(ctx.selection.get().active).toEqual({ row: 4, col: 2 });
	});

	it('warns about an invalid reference', () => {
		const { ctx, nameInput } = mount();
		typeName(nameInput, '1abc!!');
		expect(ctx.toasts).toEqual(["Reference isn't valid."]);
	});

	it('defines a new name for the selection', () => {
		const { ctx, nameInput } = mount();
		ctx.selection.set({ ranges: [{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }] });
		typeName(nameInput, 'MyRange');
		expect(ctx.workbook()!.definedNames.find((d) => d.name === 'MyRange')?.formula).toBe(
			'Sheet1!$A$1:$B$2',
		);
	});

	it('does not define names in read-only mode', () => {
		const { ctx, nameInput } = mount();
		ctx.setReadOnly(true);
		typeName(nameInput, 'MyRange');
		expect(ctx.workbook()!.definedNames).toHaveLength(1);
		expect(ctx.toasts).toHaveLength(1);
	});

	it('resolves typed text', () => {
		const wb = twoSheets();
		expect(resolveNameBox(wb, 0, 'A:C')).toMatchObject({ kind: 'go', sheet: 0 });
		expect(resolveNameBox(wb, 0, 'New_Name')).toEqual({ kind: 'define', name: 'New_Name' });
		expect(resolveNameBox(wb, 0, 'Nope!A1')).toEqual({ kind: 'invalid' });
		expect(
			absoluteReference('My Sheet', { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }),
		).toBe("'My Sheet'!$A$1");
	});
});

describe('formula input', () => {
	it('shows the active cell content', () => {
		const { ctx, formula } = mount();
		expect(formula.value).toBe('=2+3');
		ctx.selection.set({ active: { row: 1, col: 0 } });
		expect(formula.value).toBe('');
	});

	it('begins editing on focus and forwards typing to the bridge', () => {
		const { ctx, formula, q } = mount();
		const { calls, bridge } = fakeDriver(ctx);
		formula.dispatchEvent(new Event('focus'));
		expect(calls[0]).toBe('begin:edit:bar:=2+3');
		expect(q('.xfb').classList.contains('xfb-editing')).toBe(true);
		formula.value = '=2+4';
		formula.dispatchEvent(new Event('input'));
		expect(bridge.state().text).toBe('=2+4');
		expect(calls).toContain('changed:bar:=2+4');
		formula.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
		expect(calls).toContain('commit:down:false');
	});

	it('follows text typed in the cell editor', () => {
		const { ctx, formula } = mount();
		const { bridge } = fakeDriver(ctx);
		bridge.begin({ mode: 'enter', source: 'cell', text: '' });
		bridge.update('=SU', 3, 'cell');
		expect(formula.value).toBe('=SU');
	});

	it('maps Tab, Ctrl+Enter and Escape', () => {
		const { ctx, formula } = mount();
		const { calls } = fakeDriver(ctx);
		const key = (init: KeyboardEventInit) => {
			formula.dispatchEvent(new Event('focus'));
			formula.dispatchEvent(new KeyboardEvent('keydown', init));
		};
		key({ key: 'Tab', shiftKey: true });
		key({ key: 'Enter', ctrlKey: true });
		key({ key: 'Escape' });
		expect(calls.filter((c) => !c.startsWith('begin'))).toEqual([
			'commit:left:false',
			'commit:none:true',
			'cancel',
		]);
	});

	it('does not edit in read-only mode', () => {
		const { ctx, formula, q } = mount();
		ctx.setReadOnly(true);
		ctx.notify();
		const { calls } = fakeDriver(ctx);
		formula.dispatchEvent(new Event('focus'));
		expect(calls).toEqual([]);
		expect(formula.readOnly).toBe(true);
		expect(q<HTMLButtonElement>('.xfb-fx').disabled).toBe(true);
	});

	it('inserts the function picked in the insert-function dialog', async () => {
		const { ctx, formula, q } = mount();
		const { calls } = fakeDriver(ctx);
		const open = vi.fn(async () => 'sum');
		ctx.dialogs.register('insert-function', open);
		q<HTMLButtonElement>('.xfb-fx').click();
		await vi.waitFor(() => expect(calls).toContain('begin:edit:bar:=SUM('));
		expect(open).toHaveBeenCalledWith(ctx, { sheet: 0, address: { row: 0, col: 0 } });
		expect(formula.value).toBe('=SUM(');
	});

	it('starts a formula when no insert-function dialog exists', async () => {
		const { ctx, q } = mount();
		const { calls } = fakeDriver(ctx);
		q<HTMLButtonElement>('.xfb-fx').click();
		await vi.waitFor(() => expect(calls).toContain('begin:edit:bar:='));
	});

	it('toggles the expanded multi-line view', () => {
		const { q } = mount();
		const expand = q<HTMLButtonElement>('.xfb-expand');
		expand.click();
		expect(q('.xfb').classList.contains('xfb-expanded')).toBe(true);
		expect(expand.getAttribute('aria-label')).toBe('Collapse Formula Bar');
		expand.click();
		expect(expand.getAttribute('aria-expanded')).toBe('false');
	});

	it('removes its DOM on dispose', () => {
		const { ctx } = mount();
		disposers.pop()?.();
		expect(ctx.root.querySelector('.xfb')).toBeNull();
	});
});
