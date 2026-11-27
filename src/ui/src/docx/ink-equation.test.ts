// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createDocument, type Paragraph } from 'ooxml-core/docx';
import { docToModel, modelToDoc } from './model-adapter';
import { insertEquation, insertEquationAfterSelection } from './equation-commands';
import { syncEquationTab } from './equation-session';
import { buildEquationPanel } from './equation-tab';

function editor(text = '') {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text }] }];
	const doc = modelToDoc(model);
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, selection: TextSelection.create(doc, 1 + text.length) }),
	});
	return { model, view };
}

describe('Ink to Math and Ink Equation in Word', () => {
	it('inserts the recognized equation after the selected equation, never over it', () => {
		const { model, view } = editor();
		expect(insertEquation(view, 'a+b')).toBe('done');
		view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, 1)));
		expect(insertEquationAfterSelection(view, '\frac{1}{2}')).toBe(true);
		const runs = (docToModel(view.state.doc, model).blocks[0] as Paragraph).runs.filter(
			(run) => run.equation,
		);
		expect(runs).toHaveLength(2);
		expect(view.state.selection).toBeInstanceOf(NodeSelection);
		expect(insertEquationAfterSelection(view, '}')).toBe(false);
	});

	it('enables Ink Equation only while a recognizer is set, with its reason otherwise', () => {
		const panel = document.createElement('div');
		panel.dataset.ribbonTab = 'equation';
		// The ribbon element's tab API, as far as the Equation tab uses it.
		const toolbar = Object.assign(document.createElement('div'), { setTabHidden: () => undefined });
		toolbar.append(panel);
		buildEquationPanel(panel);
		const ink = panel.querySelector<HTMLButtonElement>('[data-equation-ink]')!;
		expect(ink.disabled).toBe(true);
		const { view } = editor();
		insertEquation(view, 'x');
		view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, 1)));
		const status = { format: 'unicode' as const, normalText: false, canEdit: true };
		syncEquationTab(toolbar, view.state, true, { ...status, inkUnavailable: 'No recognizer' });
		expect(ink.disabled).toBe(true);
		expect(ink.title).toBe('No recognizer');
		syncEquationTab(toolbar, view.state, true, { ...status, inkTitle: 'Ink Equation' });
		expect(ink.disabled).toBe(false);
		expect(ink.title).toBe('Ink Equation');
	});
});
