// @vitest-environment jsdom
import { createDocument, type DocumentModel } from 'ooxml-core/docx';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it } from 'vitest';
import { schema } from './schema';
import { mountReadNotes, readNoteReferences } from './view-read-notes';

afterEach(() => document.body.replaceChildren());

const model: DocumentModel = {
	...createDocument(),
	footnotes: [
		{ id: '1', blocks: [{ type: 'paragraph', id: 'n1', runs: [{ text: 'The footnote text' }] }] },
	],
	endnotes: [
		{ id: '2', blocks: [{ type: 'paragraph', id: 'n2', runs: [{ text: 'The endnote text' }] }] },
	],
};

function setup() {
	const frame = document.createElement('div');
	document.body.append(frame);
	const doc = schema.node('doc', null, [
		schema.node('paragraph', { id: 'p' }, [
			schema.text('Claim'),
			schema.node('noteReference', { kind: 'footnote', id: '1', number: 1 }),
			schema.text(' and more'),
			schema.node('noteReference', { kind: 'endnote', id: '2', number: 1, label: 'i' }),
		]),
	]);
	const view = new EditorView(frame, {
		state: EditorState.create({ doc }),
		editable: () => false,
	});
	const notes = mountReadNotes(
		frame,
		() => 'en',
		() => model,
	);
	notes.sync(view);
	const references = [...frame.querySelectorAll<HTMLElement>('[data-read-note]')];
	return { frame, view, notes, references };
}

describe('Read Mode notes', () => {
	it('lists note references in document order', () => {
		const { view } = setup();
		expect(readNoteReferences(view.state.doc).map(({ kind, label }) => [kind, label])).toEqual([
			['footnote', '1'],
			['endnote', 'i'],
		]);
	});

	it('makes each reference a keyboard button named after its note', () => {
		const { references } = setup();
		expect(references.map((element) => element.getAttribute('aria-label'))).toEqual([
			'Footnote 1',
			'Endnote i',
		]);
		for (const element of references) {
			expect(element.tabIndex).toBe(0);
			expect(element.getAttribute('role')).toBe('button');
			expect(element.getAttribute('aria-haspopup')).toBe('dialog');
		}
	});

	it('shows the note on click and returns to the reference with Back to text', () => {
		const { frame, references } = setup();
		references[0]!.click();
		const callout = frame.querySelector<HTMLElement>('.dve-reader-note');
		expect(callout?.getAttribute('role')).toBe('dialog');
		expect(callout?.textContent).toContain('The footnote text');
		const back = callout!.querySelector<HTMLButtonElement>('button')!;
		expect(document.activeElement).toBe(back);
		back.click();
		expect(frame.querySelector('.dve-reader-note')).toBeNull();
		expect(document.activeElement).toBe(references[0]);
	});

	it('opens with Enter and closes with Escape through the view', () => {
		const { frame, notes, references } = setup();
		references[1]!.focus();
		const event = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
		Object.defineProperty(event, 'target', { value: references[1] });
		expect(notes.keydown(event)).toBe(true);
		expect(frame.querySelector('.dve-reader-note')?.textContent).toContain('The endnote text');
		expect(notes.escape()).toBe(true);
		expect(frame.querySelector('.dve-reader-note')).toBeNull();
		expect(document.activeElement).toBe(references[1]);
		expect(notes.escape()).toBe(false);
	});

	it('removes its decorations and callout on dispose', () => {
		const { frame, notes, references, view } = setup();
		references[0]!.click();
		notes.dispose();
		expect(frame.querySelector('.dve-reader-note')).toBeNull();
		expect(view.dom.querySelector('[data-read-note]')).toBeNull();
	});
});
