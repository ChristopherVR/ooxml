// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';
import { createDocument, saveDocx } from 'ooxml-core/docx';
import { attachPictureDrop, dragHasPicture, pictureFiles } from './picture-drop';

registerDocxEditor();

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const png = (name = 'photo.png') => new File([PNG], name, { type: 'image/png' });

/** jsdom has no DataTransfer; this carries what the handlers read. */
function transfer(files: File[], text: Record<string, string> = {}): DataTransfer {
	return {
		files,
		types: [...(files.length ? ['Files'] : []), ...Object.keys(text)],
		items: files.map((file) => ({ kind: 'file', type: file.type })),
		getData: (type: string) => text[type] ?? '',
		dropEffect: 'none',
	} as unknown as DataTransfer;
}

function dragEvent(type: string, data: DataTransfer): Event {
	const event = new Event(type, { bubbles: true, cancelable: true, composed: true });
	Object.defineProperty(event, 'dataTransfer', { value: data });
	Object.assign(event, { clientX: 10, clientY: 10 });
	return event;
}

describe('picture drop and paste', () => {
	afterEach(() => document.body.replaceChildren());

	it('recognises picture files and ignores text and other files', () => {
		const text = new File(['x'], 'notes.txt', { type: 'text/plain' });
		expect(pictureFiles(transfer([text, png(), png('b.gif')]))).toHaveLength(2);
		expect(dragHasPicture(transfer([png()]))).toBe(true);
		expect(dragHasPicture(transfer([text]))).toBe(false);
		expect(dragHasPicture(transfer([], { 'text/plain': 'hello' }))).toBe(false);
	});

	it('claims a picture drop, places the caret and inserts each file in turn', async () => {
		const canvas = document.createElement('div');
		const order: string[] = [];
		const insert = vi.fn(async (file: File) => (order.push(file.name), true));
		const placeCaret = vi.fn(() => true);
		const view = {} as EditorView;
		const detach = attachPictureDrop(canvas, { editableView: () => view, placeCaret, insert });
		const over = dragEvent('dragover', transfer([png()]));
		canvas.dispatchEvent(over);
		expect(over.defaultPrevented).toBe(true);
		const outside = vi.fn();
		document.body.append(canvas);
		document.addEventListener('drop', outside);
		const drop = dragEvent('drop', transfer([png('a.png'), png('b.png')]));
		canvas.dispatchEvent(drop);
		await vi.waitFor(() => expect(order).toEqual(['a.png', 'b.png']));
		expect(drop.defaultPrevented).toBe(true);
		expect(placeCaret).toHaveBeenCalledWith(10, 10);
		expect(outside).not.toHaveBeenCalled();
		const plain = dragEvent('drop', transfer([], { 'text/plain': 'words' }));
		canvas.dispatchEvent(plain);
		expect(plain.defaultPrevented).toBe(false);
		document.removeEventListener('drop', outside);
		detach();
	});

	it('leaves a drop alone while the document is read-only', () => {
		const canvas = document.createElement('div');
		const insert = vi.fn(async () => true);
		attachPictureDrop(canvas, { editableView: () => undefined, insert });
		const drop = dragEvent('drop', transfer([png()]));
		canvas.dispatchEvent(drop);
		expect(drop.defaultPrevented).toBe(false);
		expect(insert).not.toHaveBeenCalled();
	});

	async function editor() {
		const element = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(element);
		await element.load(await saveDocx(createDocument()));
		const view = (element as unknown as { view: EditorView }).view;
		const pictures = () => {
			let count = 0;
			view.state.doc.descendants((node) => {
				if (node.type.name === 'image') count++;
			});
			return count;
		};
		return { element, view, pictures };
	}

	it('inserts a dropped picture file on the page as an inline picture', async () => {
		const { element, pictures } = await editor();
		const canvas = element.shadowRoot!.querySelector<HTMLElement>('.dve-canvas')!;
		canvas.dispatchEvent(dragEvent('drop', transfer([png()])));
		await vi.waitFor(() => expect(pictures()).toBe(1));
		const image = element.documentModel!.blocks.flatMap((block) =>
			block.type === 'paragraph' ? block.runs.flatMap((run) => (run.image ? [run.image] : [])) : [],
		)[0];
		expect(image).toMatchObject({ contentType: 'image/png', altText: 'photo' });
		expect(image?.anchored).toBeUndefined();
	});

	it('pastes a copied picture file, but keeps text that carries a bitmap as text', async () => {
		const { view, pictures } = await editor();
		const paste = (data: DataTransfer) => {
			const event = new Event('paste', { bubbles: true, cancelable: true });
			Object.defineProperty(event, 'clipboardData', { value: data });
			view.dom.dispatchEvent(event);
		};
		paste(transfer([png()], { 'text/plain': 'Copied words' }));
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(pictures()).toBe(0);
		paste(transfer([png()]));
		await vi.waitFor(() => expect(pictures()).toBe(1));
	});
});
