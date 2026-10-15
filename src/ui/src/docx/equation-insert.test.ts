// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { history, undo } from 'prosemirror-history';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createDocument, loadDocx, saveDocx, type Paragraph } from 'ooxml-core/docx';
import { convertOmmlToLatex, EQUATION_GALLERY, parseOmml } from 'ooxml-core/math';
import { docToModel, modelToDoc } from './model-adapter';
import { createEquationDialog } from './equation-dialog';
import {
	aloneOnItsLine,
	equationMarkup,
	insertEquation,
	selectedEquation,
	updateEquation,
} from './equation-commands';

function editor(text = '') {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text }] }];
	const doc = modelToDoc(model);
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, 1 + text.length),
			plugins: [history()],
		}),
	});
	return { model, view };
}
const equationRuns = (model: ReturnType<typeof docToModel>) =>
	(model.blocks[0] as Paragraph).runs.filter((run) => run.equation);

describe('Insert > Equation', () => {
	it('inserts a display equation into an empty paragraph and an inline one into text', () => {
		const empty = editor();
		expect(aloneOnItsLine(empty.view.state)).toBe(true);
		expect(insertEquation(empty.view, 'a^2+b^2=c^2')).toBe('done');
		const [display] = equationRuns(docToModel(empty.view.state.doc, empty.model));
		expect(display?.equation?.display).toBe(true);
		expect(display?.equation?.omml).toContain('<m:oMathPara');

		const text = editor('Area ');
		expect(aloneOnItsLine(text.view.state)).toBe(false);
		insertEquation(text.view, 'A=\\pi r^2');
		const [inline] = equationRuns(docToModel(text.view.state.doc, text.model));
		expect(inline?.equation?.display).toBe(false);
		expect((text.view.state.doc.firstChild?.textContent ?? '').trim()).toBe('Area');
	});

	it('refuses LaTeX it cannot convert and leaves the document alone', () => {
		const { view } = editor('x');
		const before = view.state.doc;
		expect(equationMarkup('\\frac{1', false)).toBeUndefined();
		expect(insertEquation(view, '\\frac{1')).toBe('invalid');
		expect(view.state.doc.eq(before)).toBe(true);
	});

	it('edits the selected equation as one undo step', () => {
		const { view, model } = editor('v ');
		insertEquation(view, 'x^2', false);
		const pos = view.state.selection.from - 1;
		view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
		expect(selectedEquation(view.state)).toMatchObject({ pos, display: false });
		expect(updateEquation(view, pos, '\\sqrt{y}', false)).toBe('done');
		const [run] = equationRuns(docToModel(view.state.doc, model));
		expect(convertOmmlToLatex(parseOmml(run!.equation!.omml))).toContain('sqrt');
		undo(view.state, view.dispatch);
		const [back] = equationRuns(docToModel(view.state.doc, model));
		expect(convertOmmlToLatex(parseOmml(back!.equation!.omml))).toContain('x');
	});

	it('inserts every gallery equation and saves valid Office Math', async () => {
		for (const entry of EQUATION_GALLERY) {
			const { view, model } = editor();
			expect(insertEquation(view, entry.latex)).toBe('done');
			const saved = await saveDocx(docToModel(view.state.doc, model));
			const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
			expect(xml).toContain('<m:oMathPara');
			const reopened = await loadDocx(saved);
			const run = (reopened.model.blocks[0] as Paragraph).runs.find((item) => item.equation);
			expect(run?.equation?.display).toBe(true);
			expect(convertOmmlToLatex(parseOmml(run!.equation!.omml)).length).toBeGreaterThan(3);
		}
	});
});

describe('equation dialog', () => {
	it('previews while typing, inserts on OK and edits the selected equation', () => {
		const { view } = editor('Sum ');
		const dialog = createEquationDialog(() => view);
		document.body.append(dialog.element);
		dialog.open();
		const source = dialog.element.querySelector('textarea')!;
		const ok = [...dialog.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!;
		expect(ok.disabled).toBe(true);
		source.value = '\\frac{a}{b}';
		source.dispatchEvent(new Event('input'));
		expect(dialog.element.querySelector('.dve-equation-preview math')).not.toBeNull();
		expect(ok.disabled).toBe(false);
		source.value = '\\frac{a';
		source.dispatchEvent(new Event('input'));
		expect(ok.disabled).toBe(true);
		expect(dialog.element.querySelector('[role="status"]')?.textContent).toContain('could not');
		source.value = 'E=mc^2';
		source.dispatchEvent(new Event('input'));
		ok.click();
		expect(view.state.doc.childCount).toBe(1);
		let found = 0;
		view.state.doc.descendants((node) => {
			if (node.type.name === 'equation') found++;
		});
		expect(found).toBe(1);
		dialog.element.remove();
	});
});
