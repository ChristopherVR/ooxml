import { expect, vi } from 'vitest';
import { mountNativeBlocks } from '../../../demos/xlsx/demo-vanilla/native-building-blocks';
import { createWorkbook, saveXlsx } from '../packages/core/src';
export async function nativeBuildingBlockContract(framework: string) {
	const host = document.createElement('div');
	document.body.append(host);
	const workbook = createWorkbook({ sheets: ['First', 'Second'] });
	const changed = vi.fn();
	const errors = vi.fn();
	await mountNativeBlocks(
		host,
		{ workbook, onWorkbookChange: changed, onWorkbookError: errors },
		framework,
	);
	const blocks = window.xlsxNativeBlocks;
	const snapshot = () => {
		blocks.flush?.();
		return blocks.snapshot();
	};
	try {
		await vi.waitFor(() => expect(snapshot().toolbarProps?.canEdit).toBe(true));
		expect(host.querySelectorAll('xlsx-editor')).toHaveLength(1);
		expect(host.querySelector('office-ui-ribbon')).toBeNull();
		const editor = host.querySelector('xlsx-editor')!;
		await vi.waitFor(() =>
			expect(
				[...host.querySelectorAll('button')].some((b) => b.textContent?.trim() === 'Bold'),
			).toBe(true),
		);
		const bold = [...host.querySelectorAll('button')].find(
			(button) => button.textContent?.trim() === 'Bold',
		)!;
		bold.click();
		await vi.waitFor(() => expect(snapshot().toolbarProps?.bold).toBe(true));
		expect(changed).toHaveBeenCalled();
		if (blocks.setCanvasClass) {
			blocks.setCanvasClass('updated-canvas');
			await vi.waitFor(() => expect(host.querySelector('.updated-canvas')).not.toBeNull());
			expect(host.querySelector('xlsx-editor')).toBe(editor);
			expect(blocks.handle.dirty).toBe(true);
		}
		const retained = snapshot().toolbarProps!;
		blocks.update({ workbook, canEdit: false, onWorkbookError: errors });
		await vi.waitFor(() => expect(snapshot().mode).toBe('view'));
		expect(await retained.onItalic()).toBe(false);
		expect(editor.readOnly).toBe(true);
		blocks.update({ workbook, canEdit: true, onWorkbookError: errors });
		await vi.waitFor(() => expect(snapshot().toolbarProps?.canEdit).toBe(true));
		await expect(blocks.handle.load(new Uint8Array([0x50, 0x4b, 3, 4]))).rejects.toThrow(/./);
		await vi.waitFor(() => expect(snapshot().error).not.toBeNull());
		const bytes = await saveXlsx(createWorkbook({ sheets: ['Replacement'] }));
		blocks.update({ bytes, onWorkbookError: errors });
		await vi.waitFor(() => expect(snapshot().toolbarProps?.sheets[0]?.name).toBe('Replacement'));
		expect(snapshot().error).toBeNull();
		expect(snapshot().loading).toBe(false);
		expect(host.querySelector('xlsx-editor')).toBe(editor);
		blocks.handle.select('B2');
		expect(snapshot().toolbarProps?.onFormulaInput('=1+2')).toBe(true);
		expect(snapshot().toolbarProps?.formula).toBe('=1+2');
		expect(snapshot().toolbarProps?.commitEdit()).toBe(true);
		const saved = await blocks.handle.saveBytes();
		await blocks.handle.load(saved);
		expect(snapshot().toolbarProps?.sheets[0]?.name).toBe('Replacement');
		await blocks.destroy();
		expect(host.querySelector('xlsx-editor')).toBeNull();
		expect(await retained.onBold()).toBe(false);
	} finally {
		if (host.querySelector('xlsx-editor')) await blocks.destroy();
		host.remove();
	}
}
