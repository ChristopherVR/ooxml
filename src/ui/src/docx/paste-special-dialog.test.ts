// @vitest-environment jsdom
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createPasteSpecialDialog, pasteSpecialFormats } from './paste-special-dialog';
import { schema } from './schema';

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
	vi.unstubAllGlobals();
	Reflect.deleteProperty(navigator, 'clipboard');
});

function setup(read: () => Promise<unknown>) {
	vi.stubGlobal('ClipboardEvent', Event);
	Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { read } });
	const host = document.body.appendChild(document.createElement('div'));
	const view = new EditorView(host, {
		state: EditorState.create({
			doc: schema.nodes.doc.create(null, schema.nodes.paragraph.create()),
		}),
		handleScrollToSelection: () => true,
	});
	views.push(view);
	const dialog = createPasteSpecialDialog(() => view);
	document.body.append(dialog.element);
	return { view, dialog };
}

const item = (types: Record<string, string>) => ({
	types: Object.keys(types),
	getType: async (type: string) => ({ text: async () => types[type] ?? '' }),
});
const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

describe('Paste Special', () => {
	it('offers the formats the clipboard holds', () => {
		expect(pasteSpecialFormats({ html: '<b>x</b>', text: 'x' }).map((f) => f.label)).toEqual([
			'HTML Format',
			'Unformatted Text',
		]);
		expect(pasteSpecialFormats({ html: null, text: 'x' }).map((f) => f.option)).toEqual(['text']);
	});

	it('pastes the chosen format', async () => {
		const { view, dialog } = setup(async () => [
			item({ 'text/html': '<p><b>Bold</b></p>', 'text/plain': 'Plain' }),
		]);
		dialog.open();
		await settle();
		await settle();
		const select = dialog.element.querySelector('select')!;
		expect([...select.options].map((option) => option.textContent)).toEqual([
			'HTML Format',
			'Unformatted Text',
		]);
		select.value = 'text';
		[...dialog.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!.click();
		expect(view.state.doc.textContent).toBe('Plain');
		expect(dialog.isOpen).toBe(false);
	});

	it('says so when the browser will not read the clipboard', async () => {
		const { view, dialog } = setup(async () => {
			throw new DOMException('denied', 'NotAllowedError');
		});
		dialog.open();
		await settle();
		await settle();
		expect(dialog.element.textContent).toContain('did not let the editor read the clipboard');
		expect(dialog.element.querySelector('select')).toBeNull();
		expect(view.state.doc.textContent).toBe('');
	});
});
