// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument, type FormField, type TextRun } from 'ooxml-core/docx';
import { formFieldOfNode } from 'ooxml-core/docx/ui';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();

const field = (formField: FormField): TextRun[] => [
	{ text: '', fieldChar: 'begin', formField },
	{ text: '', fieldCode: formField.kind === 'checkbox' ? ' FORMCHECKBOX ' : ' FORMDROPDOWN ' },
	{ text: '', fieldChar: 'end' },
];

/** A form with a check box and a drop-down, protected for filling in unless `open`. */
function mount(open = false) {
	const model = createDocument();
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p1',
			runs: [
				{ text: 'Approved ' },
				...field({ kind: 'checkbox', name: 'Approved', checkBox: { default: false } }),
				{ text: ' Status ' },
				...field({ kind: 'dropDown', name: 'Status', dropDown: { entries: ['New', 'Revised'] } }),
			],
		},
	];
	if (!open) model.protection = { edit: 'forms', enforced: true };
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	return { editor, view: (editor as unknown as { view: EditorView }).view };
}

function markerOf(view: EditorView, kind: FormField['kind']): number {
	let found = -1;
	view.state.doc.descendants((node, pos) => {
		if (formFieldOfNode(node)?.kind === kind) found = pos;
	});
	return found;
}

/** Puts the caret at `pos` and presses `key` the way the editor receives it. */
function press(view: EditorView, pos: number, key: string, init: KeyboardEventInit = {}): void {
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
	const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init });
	view.someProp('handleKeyDown', (handle) => handle(view, event));
}

/** The open list of the drop-down's entries, if any. */
const entriesList = (editor: DocxEditorElement): Element | undefined =>
	[...(editor.shadowRoot?.querySelectorAll('[role="listbox"]') ?? [])].find(
		(list) => list.textContent === 'NewRevised',
	);

const closeList = () =>
	document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

const checked = (view: EditorView) =>
	Boolean(formFieldOfNode(view.state.doc.nodeAt(markerOf(view, 'checkbox')))?.checkBox?.checked);

describe('legacy form fields from the keyboard', () => {
	afterEach(() => document.body.replaceChildren());

	it('toggles a check box with Space or Enter beside it in a protected form', () => {
		const { view, editor } = mount();
		const box = markerOf(view, 'checkbox');
		press(view, box, ' ');
		expect(checked(view)).toBe(true);
		press(view, box + 3, 'Enter');
		expect(checked(view)).toBe(false);
		press(view, box + 1, ' ');
		expect(checked(view)).toBe(true);
		const [paragraph] = editor.documentModel!.blocks;
		expect(
			paragraph?.type === 'paragraph' && paragraph.runs.find((run) => run.formField),
		).toMatchObject({ formField: { checkBox: { checked: true } } });
		expect(view.state.doc.childCount).toBe(1);
	});

	it('toggles a check box that Tab into its table cell selects, as Word does', () => {
		const { view } = mount();
		const box = markerOf(view, 'checkbox');
		// Selecting only the field's markers leaves the caret just after it.
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, box, box + 3)));
		expect(view.state.selection.empty).toBe(true);
		const event = new KeyboardEvent('keydown', { key: ' ', bubbles: true, cancelable: true });
		view.someProp('handleKeyDown', (handle) => handle(view, event));
		expect(checked(view)).toBe(true);
		// A selection that also holds text is typing over it, not operating the box.
		view.dispatch(
			view.state.tr.setSelection(TextSelection.create(view.state.doc, box - 2, box + 3)),
		);
		view.someProp('handleKeyDown', (handle) => handle(view, event));
		expect(checked(view)).toBe(true);
	});

	it('opens a drop-down list with Alt+Down, Space or Enter beside it', () => {
		const { view, editor } = mount();
		const list = markerOf(view, 'dropDown');
		press(view, list, 'ArrowDown', { altKey: true });
		expect(entriesList(editor)).toBeDefined();
		closeList();
		expect(entriesList(editor)).toBeUndefined();
		press(view, list + 3, ' ');
		expect(entriesList(editor)).toBeDefined();
		closeList();
		expect(entriesList(editor)).toBeUndefined();
		press(view, list + 2, 'Enter');
		expect(entriesList(editor)).toBeDefined();
		closeList();
	});

	it('picks the entry chosen when a drop-down lists one name twice', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p1',
				runs: field({
					kind: 'dropDown',
					name: 'Grade',
					dropDown: { entries: ['Pass', 'Fail', 'Pass'], result: 0 },
				}),
			},
		];
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.documentModel = model;
		document.body.append(editor);
		const view = (editor as unknown as { view: EditorView }).view;
		press(view, markerOf(view, 'dropDown'), 'ArrowDown', { altKey: true });
		const options = [
			...(editor.shadowRoot?.querySelectorAll<HTMLButtonElement>('[role="option"]') ?? []),
		];
		expect(options.map((option) => option.textContent)).toEqual(['Pass', 'Fail', 'Pass']);
		expect(options.map((option) => option.getAttribute('aria-selected'))).toEqual([
			'true',
			'false',
			'false',
		]);
		// The editor takes focus back after a pick, which jsdom cannot do inside a shadow root.
		const focus = vi.spyOn(view, 'focus').mockImplementation(() => {});
		options[2]!.click();
		focus.mockRestore();
		const result = formFieldOfNode(view.state.doc.nodeAt(markerOf(view, 'dropDown')));
		expect(result?.dropDown?.result).toBe(2);
	});

	it('leaves Space and Enter to typing where they can edit, as in an unprotected form', () => {
		const { view, editor } = mount(true);
		const box = markerOf(view, 'checkbox');
		press(view, box + 3, ' ');
		expect(checked(view)).toBe(false);
		press(view, box, 'Enter');
		expect(checked(view)).toBe(false);
		expect(view.state.doc.childCount).toBe(2);
		// Alt+Down has no edit to make, so it still opens a drop-down's list.
		press(view, markerOf(view, 'dropDown'), 'ArrowDown', { altKey: true });
		expect(entriesList(editor)).toBeDefined();
		closeList();
	});

	it('leaves Space to the caret the browser shows when the editor has not read it yet', () => {
		const { view } = mount();
		const box = markerOf(view, 'checkbox');
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, box)));
		// A click put the browser's caret in the label; the editor still has the old position.
		const label = view.dom.querySelector('p span')!.firstChild!;
		const root = view.root as ShadowRoot & { getSelection(): Selection | null };
		const browser = vi
			.spyOn(root, 'getSelection')
			.mockReturnValue({ focusNode: label, focusOffset: 2 } as unknown as Selection);
		try {
			const event = new KeyboardEvent('keydown', { key: ' ', bubbles: true, cancelable: true });
			view.someProp('handleKeyDown', (handle) => handle(view, event));
			expect(checked(view)).toBe(false);
		} finally {
			browser.mockRestore();
		}
	});

	it('leaves keys alone away from a check box or drop-down, or with a modifier', () => {
		const { view } = mount();
		const box = markerOf(view, 'checkbox');
		press(view, box - 1, ' ');
		press(view, box, ' ', { ctrlKey: true });
		press(view, box, ' ', { shiftKey: true });
		press(view, box, 'ArrowDown', { altKey: true });
		expect(checked(view)).toBe(false);
	});
});
