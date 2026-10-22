// @vitest-environment jsdom
import { createWorkbook, getCell, type DataValidation } from 'ooxml-core/xlsx';
import { editBridge } from 'ooxml-core/xlsx/ui';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { routeValidationAlerts } from './alert-route';
import { mountGrid } from './index';
import { createTestContext } from './test-context';
import type { ValidationAlertRequest } from '../surface-actions';

let dispose: (() => void) | undefined;
afterEach(() => {
	dispose?.();
	dispose = undefined;
	document.body.replaceChildren();
});

type Style = 'stop' | 'warning' | 'information';
function setup(style: Style, open?: (request: ValidationAlertRequest) => void, surface = true) {
	const wb = createWorkbook({ sheets: ['Data'] });
	const rule: DataValidation = {
		ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
		type: 'whole',
		operator: 'between',
		formula1: '1',
		formula2: '10',
		errorStyle: style,
		showErrorMessage: true,
		error: 'Between 1 and 10',
		errorTitle: 'Range',
	};
	wb.sheets[0]!.dataValidations.push(rule);
	const ctx = createTestContext(wb);
	if (surface) routeValidationAlerts(ctx, () => open);
	const container = document.createElement('div');
	ctx.root.append(container);
	dispose = mountGrid(ctx, container);
	const sink = ctx.root.querySelector<HTMLTextAreaElement>('.xg-cell-field textarea')!;
	sink.value = '50';
	sink.setSelectionRange(2, 2);
	sink.dispatchEvent(new Event('input', { bubbles: true }));
	sink.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
	);
	const value = () => getCell(ctx.workbook()!.sheets[0]!, 0, 0)?.value;
	return { ctx, value, builtIn: () => ctx.root.querySelector('[role="alertdialog"]') };
}
const flush = () => Promise.resolve().then(() => Promise.resolve());

describe('host-owned validation alert', () => {
	it('describes the refused edit and draws no built-in alert', () => {
		let seen: ValidationAlertRequest | undefined;
		const { builtIn } = setup('stop', (request) => (seen = request));
		expect(builtIn()).toBeNull();
		expect(seen).toMatchObject({
			ref: 'A1',
			sheet: 0,
			sheetName: 'Data',
			value: '50',
			style: 'stop',
			title: 'Range',
			message: 'Between 1 and 10',
		});
		expect(typeof seen!.cancel).toBe('function');
		expect('accept' in seen!).toBe(false);
	});

	it('stop: retry keeps the editor open with the draft', async () => {
		let seen: ValidationAlertRequest | undefined;
		const { ctx, value } = setup('stop', (request) => (seen = request));
		if (seen?.style !== 'stop') throw new Error('expected a stop alert');
		seen.retry();
		await flush();
		expect(ctx.grid()?.isEditing()).toBe(true);
		expect(editBridge(ctx).state().text).toBe('50');
		expect(value()).toBeUndefined();
	});

	it('stop: cancel restores the old value and ends the edit', async () => {
		let seen: ValidationAlertRequest | undefined;
		const { ctx, value } = setup('stop', (request) => (seen = request));
		seen!.cancel();
		await flush();
		expect(ctx.grid()?.isEditing()).toBe(false);
		expect(value()).toBeUndefined();
	});

	it('warning: accept commits (Yes), retry keeps editing (No), cancel ends', async () => {
		for (const choice of ['accept', 'retry', 'cancel'] as const) {
			let seen: ValidationAlertRequest | undefined;
			const { ctx, value } = setup('warning', (request) => (seen = request));
			(seen as unknown as Record<string, () => void>)[choice]!();
			await flush();
			expect(ctx.grid()?.isEditing()).toBe(choice === 'retry');
			expect(value()).toBe(choice === 'accept' ? 50 : undefined);
			dispose?.();
			dispose = undefined;
			document.body.replaceChildren();
		}
	});

	it('information: accept commits (OK), cancel ends', async () => {
		for (const choice of ['accept', 'cancel'] as const) {
			let seen: ValidationAlertRequest | undefined;
			const { ctx, value } = setup('information', (request) => (seen = request));
			expect(seen!.style).toBe('information');
			expect('retry' in seen!).toBe(false);
			(seen as unknown as Record<string, () => void>)[choice]!();
			await flush();
			expect(ctx.grid()?.isEditing()).toBe(false);
			expect(value()).toBe(choice === 'accept' ? 50 : undefined);
			dispose?.();
			dispose = undefined;
			document.body.replaceChildren();
		}
	});

	it('resolving twice is a no-op', async () => {
		let seen: ValidationAlertRequest | undefined;
		const { ctx, value } = setup('warning', (request) => (seen = request));
		if (seen?.style !== 'warning') throw new Error('expected a warning alert');
		seen.retry();
		seen.accept();
		seen.cancel();
		await flush();
		expect(ctx.grid()?.isEditing()).toBe(true);
		expect(value()).toBeUndefined();
	});

	it('without a callback every style acts like Retry and draws nothing', () => {
		for (const style of ['stop', 'warning', 'information'] as const) {
			const { ctx, value, builtIn } = setup(style, undefined);
			expect(builtIn()).toBeNull();
			expect(ctx.grid()?.isEditing()).toBe(true);
			expect(value()).toBeUndefined();
			expect(ctx.events.filter((e) => e.type === 'validation-error')).toHaveLength(1);
			dispose?.();
			dispose = undefined;
			document.body.replaceChildren();
		}
	});

	it('a throwing callback acts like Retry', async () => {
		const open = vi.fn(() => {
			throw new Error('host failed');
		});
		const { ctx } = setup('warning', open);
		await flush();
		expect(open).toHaveBeenCalledTimes(1);
		expect(ctx.grid()?.isEditing()).toBe(true);
	});

	it('the full editor (no surface route) keeps its built-in alert', () => {
		const { builtIn } = setup('stop', undefined, false);
		expect(builtIn()).not.toBeNull();
	});
});
