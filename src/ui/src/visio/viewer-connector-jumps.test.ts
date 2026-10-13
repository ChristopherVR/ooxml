import JSZip from 'jszip';
import { afterEach, expect, it } from 'vitest';
import { setupFormattingViewer } from './__fixtures__/formatting-viewer';

afterEach(() => document.body.replaceChildren());
it('shows connector-only jump controls and saves each change with undo and redo', async () => {
	const ui = await setupFormattingViewer();
	try {
		ui.selection();
		const field = (key: string) =>
			ui.root.querySelector<HTMLSelectElement>(`[data-pane-field="${key}"]`)!;
		expect(field('lineJumpCode').closest('label')!.hidden).toBe(true);
		const zip = await JSZip.loadAsync(ui.bytes);
		const xml = await zip.file('visio/pages/page1.xml')!.async('string');
		zip.file('visio/pages/page1.xml', xml.replace('<Text>', '<Cell N="OneD" V="1"/><Text>'));
		await ui.controller.load(await zip.generateAsync({ type: 'uint8array' }));
		ui.selection();
		expect(field('lineJumpCode').closest('label')!.hidden).toBe(false);
		for (const [key, value] of [
			['lineJumpCode', '2'],
			['lineJumpStyle', '3'],
			['lineJumpDirectionX', '2'],
			['lineJumpDirectionY', '1'],
		]) {
			field(key!).value = value!;
			field(key!).dispatchEvent(new Event('change', { bubbles: true }));
			await ui.done();
			expect(ui.feedback.at(-1)).toBe('Updated shape formatting.');
			expect(field(key!).value).toBe(value);
		}
		expect(ui.shape()).toMatchObject({
			lineJumpCode: 2,
			lineJumpStyle: 3,
			lineJumpDirectionX: 2,
			lineJumpDirectionY: 1,
		});
		ui.press('undo');
		await ui.done();
		expect(field('lineJumpDirectionY').value).toBe('0');
		ui.press('redo');
		await ui.done();
		expect(field('lineJumpDirectionY').value).toBe('1');
	} finally {
		ui.dispose();
		ui.controller.destroy();
	}
});
