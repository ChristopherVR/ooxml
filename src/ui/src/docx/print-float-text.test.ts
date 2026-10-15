// @vitest-environment jsdom
import { createDocument } from 'ooxml-core/docx';
import { history } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it } from 'vitest';
import { modelToDoc } from './model-adapter';
import { editShapeText } from './print-float-text';
import { insertDrawnShape, selectedShape } from './shape-commands';

const cleanup: Array<() => void> = [];
afterEach(() => {
	while (cleanup.length) cleanup.pop()?.();
});

function setup() {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Hello' }] }];
	const doc = modelToDoc(model);
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, 2),
			plugins: [history()],
		}),
	});
	insertDrawnShape(view, 'rect', 2, { x: 10, y: 5, widthPx: 100, heightPx: 50 });
	const sheet = document.createElement('div');
	sheet.className = 'dve-print-page';
	const float = document.createElement('div');
	Object.assign(float.style, { left: '10px', top: '5px', width: '100px', height: '50px' });
	sheet.append(float);
	document.body.append(sheet);
	cleanup.push(() => {
		view.destroy();
		host.remove();
		sheet.remove();
	});
	return { view, float, sheet };
}

describe('editing a floating shape text on the page', () => {
	it('lays a field over the shape and applies its lines as one change', () => {
		const { view, float, sheet } = setup();
		let closed = 0;
		const close = editShapeText(view, float, [], () => closed++)!;
		const field = sheet.querySelector<HTMLTextAreaElement>('.dve-print-float-text')!;
		expect(field.style.width).toBe('100px');
		field.value = 'one\ntwo';
		field.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true }));
		expect(sheet.querySelector('.dve-print-float-text')).toBeNull();
		expect(selectedShape(view.state)?.spec.text).toEqual(['one', 'two']);
		expect(closed).toBe(1);
		close();
		expect(closed).toBe(1);
	});

	it('Escape cancels without changing the shape', () => {
		const { view, float, sheet } = setup();
		editShapeText(view, float, ['keep'], () => undefined);
		const field = sheet.querySelector<HTMLTextAreaElement>('.dve-print-float-text')!;
		field.value = 'changed';
		field.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(sheet.querySelector('.dve-print-float-text')).toBeNull();
		expect(selectedShape(view.state)?.spec.text).toBeUndefined();
	});
});
