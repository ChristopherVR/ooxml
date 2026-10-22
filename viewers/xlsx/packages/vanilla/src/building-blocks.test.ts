// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { createViewerBuildingBlocks } from './index';
import { createWorkbook } from '../../core/src/index';

afterEach(() => document.body.replaceChildren());

function compose(options: Parameters<typeof createViewerBuildingBlocks>[0] = {}) {
	const blocks = createViewerBuildingBlocks({
		workbook: createWorkbook({ sheets: ['One', 'Two'] }),
		canEdit: true,
		...options,
	});
	const [toolbar, formula, tabs, canvas] = [1, 2, 3, 4].map(() => {
		const el = document.createElement('div');
		document.body.append(el);
		return el;
	});
	const stops = [
		blocks.mountToolbar(toolbar!),
		blocks.mountFormulaBar(formula!),
		blocks.mountSheetTabs(tabs!),
		blocks.mountCanvas(canvas!),
	];
	return { blocks, toolbar: toolbar!, formula: formula!, tabs: tabs!, canvas: canvas!, stops };
}
const button = (host: HTMLElement, name: string) =>
	[...host.querySelectorAll('button')].find((b) => b.textContent === name)!;

it('mounts one engine with a canvas and optional controls driven by its state', async () => {
	const { blocks, toolbar, formula, tabs, canvas } = compose();
	expect(canvas.querySelectorAll('xlsx-editor')).toHaveLength(1);
	expect(canvas.querySelector('office-ui-ribbon, office-ui-title-bar')).toBeNull();
	expect(blocks.getState().toolbarProps?.canEdit).toBe(true);
	expect([...toolbar.querySelectorAll('button')].map((b) => b.textContent)).toEqual([
		'Undo',
		'Redo',
		'Bold',
		'Italic',
		'Underline',
		'Align Left',
		'Center',
		'Align Right',
	]);
	expect([...toolbar.querySelectorAll('select')].map((s) => s.getAttribute('aria-label'))).toEqual([
		'Font',
		'Font Size',
		'Number Format',
	]);
	expect(toolbar.getAttribute('role')).toBe('toolbar');
	expect(formula.querySelectorAll('input')).toHaveLength(2);
	expect([...tabs.querySelectorAll('[role="tab"]')].map((t) => t.textContent)).toEqual([
		'One',
		'Two',
	]);
	button(toolbar, 'Bold').click();
	await vi.waitFor(() => expect(blocks.getState().toolbarProps?.bold).toBe(true));
	expect(button(toolbar, 'Bold').getAttribute('aria-pressed')).toBe('true');
	expect(blocks.getState().dirty).toBe(true);
	blocks.destroy();
});

it('keeps the focused control and the typed draft while state updates', async () => {
	const { blocks, toolbar, formula, tabs } = compose();
	const formulaInput = formula.querySelectorAll('input')[1]!;
	const bold = button(toolbar, 'Bold');
	formulaInput.focus();
	formulaInput.value = '=1+2';
	formulaInput.dispatchEvent(new Event('input', { bubbles: true }));
	await vi.waitFor(() => expect(blocks.getState().toolbarProps?.editing).toBe(true));
	// A toolbar click does not take focus from the formula bar.
	const down = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
	bold.dispatchEvent(down);
	expect(down.defaultPrevented).toBe(true);
	// A state refresh patches the controls in place: same elements, focus and draft intact.
	blocks.update({ canEdit: true });
	await Promise.resolve();
	expect(formula.querySelectorAll('input')[1]).toBe(formulaInput);
	expect(button(toolbar, 'Bold')).toBe(bold);
	expect(document.activeElement).toBe(formulaInput);
	expect(formulaInput.value).toBe('=1+2');
	blocks.getState().toolbarProps!.onCancelEdit();
	formulaInput.focus();
	await blocks.getState().toolbarProps!.onUnderline();
	await vi.waitFor(() => expect(blocks.getState().toolbarProps?.underline).toBe(true));
	expect(button(toolbar, 'Underline').getAttribute('aria-pressed')).toBe('true');
	expect(formula.querySelectorAll('input')[1]).toBe(formulaInput);
	expect(document.activeElement).toBe(formulaInput);
	// Sheet tabs patch in place too.
	const first = tabs.querySelectorAll('[role="tab"]')[0]!;
	(tabs.querySelectorAll('[role="tab"]')[1] as HTMLElement).click();
	await vi.waitFor(() => expect(blocks.getState().toolbarProps?.activeSheet).toBe(1));
	expect(tabs.querySelectorAll('[role="tab"]')[0]).toBe(first);
	expect(first.getAttribute('aria-selected')).toBe('false');
	blocks.destroy();
});

it('selects from the name box and flags an invalid reference', async () => {
	const { blocks, formula } = compose();
	const name = formula.querySelectorAll('input')[0]!;
	name.value = 'C3';
	name.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
	expect(blocks.handle.getSelection()).toBe('C3');
	expect(name.getAttribute('aria-invalid')).toBe('false');
	name.value = 'not a ref';
	name.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
	expect(name.getAttribute('aria-invalid')).toBe('true');
	blocks.destroy();
});

it('disables controls in view mode and routes dialog workflows to host callbacks', async () => {
	const onOpenShareDialog = vi.fn();
	const { blocks, toolbar } = compose({ canEdit: false, onOpenShareDialog });
	expect(blocks.getState().mode).toBe('view');
	expect(button(toolbar, 'Bold').disabled).toBe(true);
	expect(toolbar.querySelector<HTMLSelectElement>('select')!.disabled).toBe(true);
	expect(blocks.getState().toolbarProps?.enabled.openShareDialog).toBe(true);
	blocks.handle.share();
	expect(onOpenShareDialog).toHaveBeenCalledTimes(1);
	blocks.update({ canEdit: true, workbook: blocks.handle.element.workbook! });
	await vi.waitFor(() => expect(button(toolbar, 'Bold').disabled).toBe(false));
	expect(blocks.getState().toolbarProps?.enabled.openShareDialog).toBe(false);
	blocks.handle.share();
	expect(onOpenShareDialog).toHaveBeenCalledTimes(1);
	blocks.destroy();
});

it('reports validation failures and clears them when the draft is canceled', async () => {
	const workbook = createWorkbook();
	workbook.sheets[0]!.dataValidations = [
		{
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '2',
			showErrorMessage: true,
			errorStyle: 'stop',
		},
	];
	const onValidationError = vi.fn();
	const { blocks, formula } = compose({ workbook, onValidationError });
	const formulaInput = formula.querySelectorAll('input')[1]!;
	formulaInput.value = '9';
	formulaInput.dispatchEvent(new Event('input', { bubbles: true }));
	formulaInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
	expect(onValidationError).toHaveBeenCalledTimes(1);
	await vi.waitFor(() => expect(blocks.getState().editError).toBeTruthy());
	expect(formulaInput.getAttribute('aria-invalid')).toBe('true');
	expect(formula.querySelector('[role="alert"]')?.textContent).toBe(blocks.getState().editError);
	formulaInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
	await vi.waitFor(() => expect(blocks.getState().editError).toBeNull());
	expect(formula.querySelector('[role="alert"]')).toBeNull();
	blocks.destroy();
});

it('hands the validation alert to the host instead of drawing one', async () => {
	const workbook = createWorkbook();
	workbook.sheets[0]!.dataValidations = [
		{
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '2',
			showErrorMessage: true,
			errorStyle: 'warning',
			error: 'One or two',
		},
	];
	const onOpenValidationAlert = vi.fn();
	const { blocks, formula, canvas } = compose({ workbook, onOpenValidationAlert });
	await vi.waitFor(() =>
		expect(blocks.getState().toolbarProps?.enabled.validationAlert).toBe(true),
	);
	const input = formula.querySelectorAll('input')[1]!;
	input.value = '9';
	input.dispatchEvent(new Event('input', { bubbles: true }));
	input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
	expect(onOpenValidationAlert).toHaveBeenCalledTimes(1);
	const request = onOpenValidationAlert.mock.calls[0]![0];
	expect(request).toMatchObject({ style: 'warning', ref: 'A1', value: '9', message: 'One or two' });
	expect(
		canvas.querySelector('xlsx-editor')!.shadowRoot!.querySelector('[role="alertdialog"]'),
	).toBeNull();
	request.retry();
	request.accept();
	await Promise.resolve();
	expect(blocks.getState().toolbarProps?.editing).toBe(true);
	blocks.destroy();
});

it('removes its controls and the engine on dispose', () => {
	const { blocks, toolbar, canvas, stops } = compose();
	for (const stop of stops) stop();
	expect(canvas.querySelector('xlsx-editor')).toBeNull();
	expect(toolbar.children).toHaveLength(0);
	blocks.destroy();
	expect(() => blocks.handle.getSelection()).toThrow('not mounted');
});

it('is read-only until the host passes canEdit: true', async () => {
	const { blocks } = compose({ canEdit: undefined });
	expect(blocks.getState().mode).toBe('view');
	await vi.waitFor(() => expect(blocks.getState().toolbarProps).not.toBeNull());
	expect(blocks.getState().toolbarProps?.canEdit).toBe(false);
	blocks.update({ canEdit: true });
	await vi.waitFor(() => expect(blocks.getState().toolbarProps?.canEdit).toBe(true));
	expect(blocks.getState().mode).toBe('edit');
});
