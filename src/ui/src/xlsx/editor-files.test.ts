// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import { loadWorkbook, saveWorkbook } from 'ooxml-core/xlsx/load';
import type { GridController } from 'ooxml-core/xlsx/ui';
import { loadInto, newInto, runFileCommand, saveBytes } from './editor-files';
import { shellFixture } from './test-support/shell';
import { packagePassword } from './dialogs/package-password';
import * as fileCommands from './file-commands';
import { mountGrid } from './grid';

vi.mock('./dialogs/package-password', () => ({ packagePassword: vi.fn() }));
afterEach(() => {
	vi.resetAllMocks();
	document.body.innerHTML = '';
});

describe('encrypted workbook UI', () => {
	it('retries a wrong password, opens the workbook and reports its warnings', async () => {
		const bytes = await saveWorkbook(createWorkbook(), 'xlsx', {
			password: 'secret',
			encryption: { spinCount: 10 },
		});
		const { core, events } = shellFixture();
		core.ctx.toast = vi.fn();
		vi.mocked(packagePassword).mockResolvedValueOnce('wrong').mockResolvedValueOnce('secret');
		await loadInto(core, bytes, 'protected.xlsx');
		expect(packagePassword).toHaveBeenCalledTimes(2);
		expect(core.ctx.toast).toHaveBeenCalledWith('The password is incorrect.', 'warning');
		expect(core.fileName).toBe('protected.xlsx');
		expect(events.some((e) => e.type === 'workbook-error')).toBe(false);
		expect(core.workbook?.warnings.join(' ')).toContain('opened with a password');
		expect(core.savePassword).toBeUndefined();
	});

	it('keeps the current workbook when cancelled or superseded by a new workbook', async () => {
		const bytes = await saveWorkbook(createWorkbook(), 'xlsx', {
			password: 'secret',
			encryption: { spinCount: 10 },
		});
		const { core } = shellFixture();
		const original = core.workbook;
		vi.mocked(packagePassword).mockResolvedValueOnce(undefined);
		await loadInto(core, bytes, 'protected.xlsx');
		expect(core.workbook).toBe(original);
		vi.mocked(packagePassword).mockImplementationOnce(async () => {
			newInto(core);
			return 'secret';
		});
		await loadInto(core, bytes, 'protected.xlsx');
		expect(core.workbook).not.toBe(original);
		expect(core.fileName).toBe('Book1.xlsx');
	});

	it('encrypts Excel saves, keeps CSV exports plain and clears the password for a new workbook', async () => {
		const { core } = shellFixture();
		core.savePassword = 'secret';
		const bytes = await saveBytes(core);
		await expect(loadWorkbook(bytes)).rejects.toMatchObject({ code: 'password-required' });
		await expect(loadWorkbook(bytes, { password: 'secret' })).resolves.toHaveProperty('sheets');
		await expect(loadWorkbook(await saveBytes(core, 'csv'))).resolves.toHaveProperty(
			'format',
			'csv',
		);
		newInto(core);
		expect(core.savePassword).toBeUndefined();
	}, 30000);

	it('keeps CSV formulas as text by default', async () => {
		const { core } = shellFixture();
		await loadInto(core, new TextEncoder().encode('=1+1'), 'data.csv');
		expect(core.workbook?.sheets[0]?.rows.get(0)?.get(0)).toMatchObject({ value: '=1+1' });
		expect(core.workbook?.sheets[0]?.rows.get(0)?.get(0)?.formula).toBeUndefined();
	});
});

describe('saving a pending cell edit', () => {
	it('blocks real grid validation until the invalid edit is cancelled', async () => {
		const { core } = shellFixture();
		const session = core.ctx.session()!;
		session.setCellInput(0, 0, 0, '1');
		session.setDataValidation(
			0,
			{
				ranges: [],
				type: 'whole',
				operator: 'between',
				formula1: '1',
				formula2: '2',
				showErrorMessage: true,
				errorStyle: 'stop',
			},
			{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
		);
		const container = document.createElement('div');
		core.ctx.root.append(container);
		const dispose = mountGrid(core.ctx, container);
		try {
			const input = container.querySelector<HTMLTextAreaElement>('textarea')!;
			core.ctx.grid()!.focus();
			input.value = '9';
			input.dispatchEvent(new Event('input', { bubbles: true }));
			expect(core.ctx.grid()!.isEditing()).toBe(true);
			for (const format of ['xlsx', 'csv'] as const)
				await expect(saveBytes(core, format)).rejects.toThrow(/cell edit/);
			expect(session.workbook.sheets[0]!.rows.get(0)?.get(0)?.value).toBe(1);
			expect(core.dirty.dirty).toBe(true);
			expect(core.ctx.root.querySelectorAll('[role="alertdialog"]')).toHaveLength(1);
			core.ctx.root.querySelector<HTMLButtonElement>('[data-answer="cancel"]')!.click();
			await Promise.resolve();
			expect(core.ctx.grid()!.isEditing()).toBe(false);
			expect(new TextDecoder().decode(await saveBytes(core, 'csv')).trim()).toBe('1');
		} finally {
			dispose();
		}
	});

	it('saves an idle grid without attempting a cell commit', async () => {
		const { core } = shellFixture();
		const grid = {
			isEditing: () => false,
			commitEdit: vi.fn(() => false),
		} as unknown as GridController;
		vi.spyOn(core.ctx, 'grid').mockReturnValue(grid);
		await expect(loadWorkbook(await saveBytes(core, 'csv'))).resolves.toHaveProperty(
			'format',
			'csv',
		);
		expect(grid.commitEdit).not.toHaveBeenCalled();
	});

	it.each(['xlsx', 'csv'] as const)('rejects %s when the edit cannot commit', async (format) => {
		const { core } = shellFixture();
		const grid = {
			isEditing: () => true,
			commitEdit: vi.fn(() => false),
		} as unknown as GridController;
		vi.spyOn(core.ctx, 'grid').mockReturnValue(grid);
		core.dirty.set(true);
		await expect(saveBytes(core, format)).rejects.toThrow(/cell edit/);
		expect(grid.commitEdit).toHaveBeenCalledOnce();
		expect(core.dirty.dirty).toBe(true);
	});

	it('reports the failed save without marking the workbook saved', async () => {
		const { core, events } = shellFixture();
		vi.spyOn(core.ctx, 'grid').mockReturnValue({
			isEditing: () => true,
			commitEdit: () => false,
		} as GridController);
		core.dirty.set(true);
		const setSaveState = vi.fn();
		const download = vi.spyOn(fileCommands, 'downloadBytes');
		await runFileCommand(core, 'save', undefined, {
			chrome: { setSaveState, fileNameChanged: vi.fn() },
		});
		expect(setSaveState.mock.calls.map(([state]) => state)).toEqual(['saving', 'dirty']);
		expect(core.dirty.dirty).toBe(true);
		expect(events.filter((event) => event.type === 'workbook-error')).toHaveLength(1);
		expect(download).not.toHaveBeenCalled();
	});

	it('exports the committed value when editing succeeds', async () => {
		const { core } = shellFixture();
		vi.spyOn(core.ctx, 'grid').mockReturnValue({
			isEditing: () => true,
			commitEdit: () => {
				core.ctx.session()!.setCellInput(0, 0, 0, '3');
				return true;
			},
		} as GridController);
		const saved = await loadWorkbook(await saveBytes(core));
		expect(saved.sheets[0]!.rows.get(0)?.get(0)?.value).toBe(3);
	});
});
