// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo } from 'prosemirror-history';
import { createReviewOptionsDialog } from './review-options-dialog';
import { schema } from './schema';
import { FormatDialogs } from './format-dialogs';
import { createDocument } from 'ooxml-core/docx';

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
});
function setup() {
	let editable = true;
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({
			doc: schema.node('doc', null, schema.node('paragraph')),
			plugins: [history()],
		}),
		editable: () => editable,
	});
	views.push(view);
	const dialog = createReviewOptionsDialog(() => view);
	document.body.append(dialog.element);
	const field = (label: string) =>
		dialog.element.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`)!;
	const change = (label: string, value: boolean) => {
		const input = field(label);
		input.checked = value;
		input.dispatchEvent(new Event('change', { bubbles: true }));
	};
	const click = (label: string) =>
		[...dialog.element.querySelectorAll<HTMLButtonElement>('button')]
			.find((button) => button.textContent === label)!
			.click();
	return {
		view,
		dialog,
		field,
		change,
		click,
		lock() {
			editable = false;
			view.updateState(view.state);
		},
	};
}
describe('tracking options dialog', () => {
	it('uses the main document preferences while the active view edits a different story', () => {
		const root = setup().view;
		const story = setup().view;
		root.dispatch(root.state.tr.setDocAttribute('trackFormatting', false));
		const dialogs = new FormatDialogs({
			view: () => story,
			reviewView: () => root,
			model: () => createDocument(),
		});
		document.body.append(...dialogs.elements);
		dialogs.open('tracking');
		const element = dialogs.elements.at(-1)!;
		const input = element.querySelector<HTMLInputElement>('[aria-label="Track formatting"]')!;
		expect(input.checked).toBe(false);
		input.checked = true;
		input.dispatchEvent(new Event('change'));
		const storyBefore = story.state.doc;
		[...element.querySelectorAll<HTMLButtonElement>('button')]
			.find((button) => button.textContent === 'OK')!
			.click();
		expect(root.state.doc.attrs.trackFormatting).toBe(true);
		expect(story.state.doc).toBe(storyBefore);
	});
	it('loads preferences and applies both changes as one undo operation', () => {
		const { view, dialog, field, change, click } = setup();
		dialog.open();
		expect(field('Track formatting').checked).toBe(true);
		expect(field('Track moves').checked).toBe(true);
		change('Track formatting', false);
		change('Track moves', false);
		click('OK');
		expect(dialog.isOpen).toBe(false);
		expect(view.state.doc.attrs).toMatchObject({ trackFormatting: false, trackMoves: false });
		expect(undo(view.state, view.dispatch)).toBe(true);
		expect(view.state.doc.attrs).toMatchObject({ trackFormatting: true, trackMoves: true });
	});
	it('does not overwrite an untouched preference changed while the dialog is open', () => {
		const { view, dialog, change, click } = setup();
		dialog.open();
		change('Track formatting', false);
		view.dispatch(view.state.tr.setDocAttribute('trackMoves', false));
		click('OK');
		expect(view.state.doc.attrs).toMatchObject({ trackFormatting: false, trackMoves: false });
	});
	it('cancels without applying changes and reads fresh values when reopened', () => {
		const { view, dialog, field, change, click } = setup();
		dialog.open();
		change('Track formatting', false);
		click('Cancel');
		expect(view.state.doc.attrs.trackFormatting).toBe(true);
		dialog.open();
		expect(field('Track formatting').checked).toBe(true);
		click('OK');
		expect(undo(view.state)).toBe(false);
	});
	it('blocks opening and applying when edit authority is revoked', () => {
		const { view, dialog, change, click, lock } = setup();
		dialog.open();
		change('Track formatting', false);
		lock();
		click('OK');
		expect(view.state.doc.attrs.trackFormatting).toBe(true);
		dialog.close();
		dialog.open();
		expect(dialog.isOpen).toBe(false);
	});
	it('localizes its heading and checkboxes using the shared dialog shell', () => {
		const { dialog } = setup();
		dialog.setLocale('fr');
		dialog.open();
		expect((dialog.element as HTMLElement & { heading: string }).heading).toBe('Options de suivi');
		expect(
			dialog.element.querySelector('input[aria-label="Suivre la mise en forme"]'),
		).not.toBeNull();
	});
});
