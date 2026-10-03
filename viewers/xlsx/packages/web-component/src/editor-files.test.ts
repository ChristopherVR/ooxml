// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook } from '@christophervr/xlsx-core';
import { loadWorkbook, saveWorkbook } from '@christophervr/xlsx-core/load';
import { loadInto, newInto, saveBytes } from './editor-files';
import { shellFixture } from './test-support/shell';
import { packagePassword } from './dialogs/package-password';

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
