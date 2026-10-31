import { afterEach, expect, it } from 'vitest';
import JSZip from 'jszip';
import { setupFormattingViewer } from './__fixtures__/formatting-viewer';
import { ViewerInlineSize } from './viewer-inline-size';

afterEach(() => document.body.replaceChildren());

it('keeps a focused size draft through rendering and commits Enter once with undo', async () => {
	const ui = await setupFormattingViewer();
	const zip = await JSZip.loadAsync(ui.bytes);
	const path = 'visio/pages/page1.xml';
	zip.file(
		path,
		(await zip.file(path)!.async('string'))
			.replaceAll('N="X" V="3"/>', 'N="X" V="3" F="Width"/>')
			.replaceAll('N="Y" V="1"/>', 'N="Y" V="1" F="Height"/>'),
	);
	await ui.controller.load(await zip.generateAsync({ type: 'uint8array' }));
	const fields = new ViewerInlineSize(ui.root, ui.controller);
	const unwire = fields.wire();
	const unsubscribe = ui.controller.subscribe((state) => fields.render(state));
	try {
		ui.selection();
		const input = fields.inputs().find((input) => input.dataset.inlineSize === 'width')!;
		const before = input.value;
		input.focus();
		expect(ui.root.activeElement).toBe(input);
		input.value = '5';
		fields.render(ui.controller.state);
		expect(input.value).toBe('5');
		expect(ui.edits).toHaveLength(0);
		input.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
		);
		await ui.done();
		expect(ui.controller.state.edit.error).toBeNull();
		expect(ui.edits).toHaveLength(1);
		expect(input.value).toBe('5');
		input.dispatchEvent(new Event('change'));
		await ui.done();
		expect(ui.edits).toHaveLength(1);
		ui.press('undo');
		await ui.done();
		input.blur();
		fields.render(ui.controller.state);
		expect(input.value).toBe(before);
	} finally {
		unsubscribe();
		unwire();
		ui.dispose();
		ui.controller.destroy();
	}
});
