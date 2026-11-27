// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { history } from 'prosemirror-history';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createDocument, loadDocx, saveDocx, type Paragraph } from 'ooxml-core/docx';
import type { EquationInputFormat } from 'ooxml-core/math';
import { equationKit, loadEquationKit } from './equation-kit';
import { docToModel, modelToDoc } from './model-adapter';
import {
	convertAllEquations,
	EquationEditingBox,
	insertNewEquation,
	placeBuiltIn,
	structureById,
} from './equation-editing';
import { insertEquationXml } from './equation-commands';
import { EquationController } from './equation-controller';
import { createRibbon } from './ribbon';
import { panelOf } from './ribbon-tab-api';

const views: EditorView[] = [];
beforeAll(() => loadEquationKit());
const buildUpEquation = (...args: Parameters<ReturnType<typeof equationKit>['buildUpEquation']>) =>
	equationKit().buildUpEquation(...args);
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
});

function editor(text = '') {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text }] }];
	const doc = modelToDoc(model);
	const view = new EditorView(document.body.appendChild(document.createElement('div')), {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, 1 + text.length),
			plugins: [history()],
		}),
	});
	views.push(view);
	let format: EquationInputFormat = 'unicode';
	const box = new EquationEditingBox({
		view: () => view,
		toolbar: () => undefined,
		locale: () => 'en',
		format: () => format,
		setFormat: (next) => (format = next),
		refresh: () => {},
	});
	return { model, view, box };
}

const equations = (view: EditorView) => {
	const found: string[] = [];
	view.state.doc.descendants((node) => {
		if (node.type.name === 'equation') found.push(String(node.attrs.omml));
	});
	return found;
};

function type(box: EquationEditingBox, text: string) {
	box.input.setRangeText(text, box.input.selectionStart ?? 0, box.input.selectionEnd ?? 0, 'end');
	box.input.dispatchEvent(new Event('input'));
}

async function saved(view: EditorView, model: ReturnType<typeof createDocument>) {
	const bytes = await saveDocx(docToModel(view.state.doc, model));
	const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
	return { bytes, xml };
}

describe('Insert New Equation', () => {
	it('inserts an empty display equation on an empty line, selected for editing', () => {
		const { view } = editor();
		const pos = insertNewEquation(view)!;
		expect(view.state.selection).toBeInstanceOf(NodeSelection);
		expect(view.state.selection.from).toBe(pos);
		expect(equations(view)[0]).toMatch(/^<m:oMathPara [^>]*><m:oMath\/><\/m:oMathPara>$/);
	});
});

describe('the equation editor', () => {
	it('builds linear input up into Office Math as it is typed and saves it', async () => {
		const { view, model, box } = editor('Let ');
		const pos = insertNewEquation(view)!;
		box.sync(view, pos, true);
		expect(box.input.placeholder).toBe('Type equation here.');
		type(box, '(a+b)/c');
		// Runs are written as Word writes them: one run per stretch of text.
		expect(equations(view)[0]).toContain('<m:f><m:num><m:r><m:t>a+b</m:t></m:r>');
		const { bytes, xml } = await saved(view, model);
		expect(xml).toMatch(/<m:oMath [^>]*><m:f><m:num>/);
		const reopened = await loadDocx(bytes);
		const run = (reopened.model.blocks[0] as Paragraph).runs.find((item) => item.equation);
		expect(run?.equation?.omml).toContain('<m:f>');
	});

	it('runs math AutoCorrect on space after a control word', () => {
		const { view, box } = editor();
		box.sync(view, insertNewEquation(view)!, true);
		type(box, 'x=\\alpha');
		const space = new KeyboardEvent('keydown', { key: ' ', cancelable: true });
		box.input.dispatchEvent(space);
		expect(space.defaultPrevented).toBe(true);
		expect(box.input.value).toBe('x=α');
		expect(equations(view)[0]).toContain('<m:t>x=α</m:t>');
	});

	it('inserts a structure at the caret with its placeholders and selects the first', () => {
		const { view, box } = editor();
		box.sync(view, insertNewEquation(view)!, true);
		type(box, 'y=');
		box.insert(structureById('fraction/stacked', 'unicode')!);
		// Each placeholder is grouped, so what is typed in it stays that argument.
		expect(box.input.value).toBe('y=〖⬚〗/〖⬚〗');
		expect([box.input.selectionStart, box.input.selectionEnd]).toEqual([3, 4]);
		// Empty arguments are saved empty, as Word saves a structure waiting for input.
		expect(equations(view)[0]).toContain('<m:f><m:num/><m:den/></m:f>');
		type(box, 'a+b');
		const tab = new KeyboardEvent('keydown', { key: 'Tab', cancelable: true });
		box.input.dispatchEvent(tab);
		type(box, '2');
		expect(box.input.value).toBe('y=〖a+b〗/〖2〗');
		expect(equations(view)[0]).toContain(
			'<m:f><m:num><m:r><m:t>a+b</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>',
		);
	});

	it('puts the selection into the first placeholder of a structure inserted around it', () => {
		const { view, box } = editor();
		box.sync(view, insertNewEquation(view)!, true);
		type(box, 'x+1');
		box.input.setSelectionRange(0, 3);
		box.insert(structureById('radical/squareRoot', 'unicode')!);
		expect(box.input.value).toBe('√〖x+1〗');
		expect(equations(view)[0]).toContain('<m:e><m:r><m:t>x+1</m:t></m:r></m:e>');
	});

	it('edits an existing equation from its linear text', () => {
		const { view, box } = editor();
		const omml = buildUpEquation('x^2', 'unicode', false)!;
		const pos = insertEquationXml(view, omml, false)!;
		box.sync(view, pos, true);
		expect(box.input.value).toBe('x^2');
		box.input.setSelectionRange(3, 3);
		type(box, '+y_i');
		expect(equations(view)[0]).toContain('<m:sSub>');
	});

	it('types LaTeX in the LaTeX input format and converts the text between formats', () => {
		const { view, box } = editor();
		box.sync(view, insertNewEquation(view)!, true);
		box.setFormat('latex');
		type(box, '\\frac{1}{2}');
		expect(equations(view)[0]).toContain('<m:f>');
		box.setFormat('unicode');
		expect(box.input.value).toBe('1/2');
	});

	it('toggles Normal Text around the caret', () => {
		const { view, box } = editor();
		box.sync(view, insertNewEquation(view)!, true);
		box.toggleNormalText();
		expect(box.input.value).toBe('""');
		expect(box.normalText).toBe(true);
		type(box, 'if');
		expect(equations(view)[0]).toContain('<m:nor m:val="1"/>');
		box.toggleNormalText();
		expect(box.input.value).toBe('if');
	});

	it('converts between Professional and Linear and saves the linear text', async () => {
		const { view, model, box } = editor();
		const pos = insertEquationXml(view, buildUpEquation('a/b', 'unicode', false)!, false)!;
		box.sync(view, pos, true);
		box.setLinear(true);
		// Word's Linear: the text in runs, each build-up operator in a run of its own.
		expect(equations(view)[0]).toMatch(
			/<m:oMath [^>]*><m:r><m:t>a<\/m:t><\/m:r><m:r><m:t>\/<\/m:t><\/m:r><m:r><m:t>b<\/m:t><\/m:r><\/m:oMath>$/,
		);
		expect((await saved(view, model)).xml).toContain('<m:t>/</m:t>');
		box.setLinear(false);
		expect(equations(view)[0]).toContain('<m:f>');
		convertAllEquations(view, true);
		expect(equations(view)[0]).not.toContain('<m:f>');
		// The equation stays selected, so its editor keeps the caret.
		expect(view.state.selection).toBeInstanceOf(NodeSelection);
		convertAllEquations(view, false);
		expect(equations(view)[0]).toContain('<m:f>');
	});

	it('keeps an equation linear text cannot hold read-only, and Convert leaves it alone', () => {
		const { view, box } = editor();
		const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
		const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const mixed = `<m:oMathPara xmlns:m="${M}" xmlns:w="${W}"><m:oMath><m:r><m:t>a</m:t></m:r><m:r><w:rPr><w:b/></w:rPr><m:t>b</m:t></m:r></m:oMath></m:oMathPara>`;
		const pos = insertEquationXml(view, mixed, true)!;
		box.sync(view, pos, true);
		expect(box.input.readOnly).toBe(true);
		expect(box.editable).toBe(false);
		expect(box.element.querySelector('[role="status"]')?.textContent).toBe(
			'This equation mixes formatting its linear text cannot hold, so it is kept as it is.',
		);
		box.insert('x');
		convertAllEquations(view, true);
		expect(equations(view)[0]).toBe(mixed);
	});

	it('says an equation with tracked changes is kept until they are accepted or rejected', () => {
		const { view, box } = editor();
		const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
		const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const tracked = `<m:oMath xmlns:m="${M}" xmlns:w="${W}"><w:ins w:id="1" w:author="A"><m:r><m:t>b</m:t></m:r></w:ins></m:oMath>`;
		box.sync(view, insertEquationXml(view, tracked, false)!, true);
		expect(box.editable).toBe(false);
		expect(box.element.querySelector('[role="status"]')?.textContent).toContain(
			'Accept or reject them to edit it.',
		);
	});

	it('edits each equation of a display paragraph with several on its own', () => {
		const { view, box } = editor();
		const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
		const first = `<m:oMath><m:r><m:t>a=b</m:t></m:r></m:oMath>`;
		const two = `<m:oMathPara xmlns:m="${M}">${first}<m:oMath><m:r><m:t>c</m:t></m:r></m:oMath></m:oMathPara>`;
		const pos = insertEquationXml(view, two, true)!;
		box.sync(view, pos, true);
		expect(box.editable).toBe(true);
		expect(box.input.value).toBe('a=b');
		expect(box.element.getAttribute('aria-label')).toBe('Equation editor, Equation 1 of 2');
		// The right arrow at the end of the first equation moves to the second.
		box.input.setSelectionRange(3, 3);
		box.input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', cancelable: true }));
		expect(box.input.value).toBe('c');
		expect(box.element.getAttribute('aria-label')).toBe('Equation editor, Equation 2 of 2');
		box.input.setSelectionRange(1, 1);
		type(box, '/2');
		const out = equations(view)[0]!;
		expect(out.startsWith(`<m:oMathPara xmlns:m="${M}">${first}`)).toBe(true);
		expect(out).toContain('<m:f>');
		// The paragraph keeps both equations, and the box stays on the second one.
		expect(equationKit().equationParts(out)).toBe(2);
		box.sync(view, pos, false, true);
		expect(box.input.value).toBe('c/2');
		// Convert > All converts each equation on its own; the left arrow at the start goes back.
		box.input.setSelectionRange(0, 0);
		box.input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', cancelable: true }));
		expect(box.input.value).toBe('a=b');
		convertAllEquations(view, true);
		const linear = equations(view)[0]!;
		expect(linear).toContain(first);
		expect(linear).not.toContain('<m:f>');
	});

	it('keeps the editor on the same equation when edits elsewhere move it', () => {
		const { view, box } = editor('ab');
		const pos = insertNewEquation(view)!;
		box.sync(view, pos, true);
		type(box, 'x+1');
		box.input.setSelectionRange(1, 2);
		// A collaborator types before the equation: the selection maps along with it.
		view.dispatch(view.state.tr.insertText('zz', 1));
		expect(view.state.selection.from).toBe(pos + 2);
		box.sync(view, pos + 2, false, true);
		expect(box.input.value).toBe('x+1');
		expect([box.input.selectionStart, box.input.selectionEnd]).toEqual([1, 2]);
	});

	it('ignores Enter and Space while an input method is composing', () => {
		const { view, box } = editor();
		box.sync(view, insertNewEquation(view)!, true);
		type(box, '\\alpha');
		const space = new KeyboardEvent('keydown', { key: ' ', isComposing: true, cancelable: true });
		box.input.dispatchEvent(space);
		expect(space.defaultPrevented).toBe(false);
		const enter = new KeyboardEvent('keydown', {
			key: 'Enter',
			isComposing: true,
			cancelable: true,
		});
		box.input.dispatchEvent(enter);
		expect(view.state.selection).toBeInstanceOf(NodeSelection);
	});

	it('does not double the brackets of a Bracket structure around a selection', () => {
		const { view, box } = editor();
		box.sync(view, insertNewEquation(view)!, true);
		type(box, 'a+b');
		box.input.setSelectionRange(0, 3);
		box.insert(structureById('bracket/parentheses', 'unicode')!);
		expect(box.input.value).toBe('(a+b)');
	});

	it('leaves the equation with the arrow keys and deletes an empty one with Backspace', () => {
		const { view, box } = editor('ab');
		const pos = insertNewEquation(view)!;
		box.sync(view, pos, true);
		box.input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Backspace', cancelable: true }));
		expect(equations(view)).toHaveLength(0);
		const again = insertNewEquation(view)!;
		box.sync(view, again, true);
		type(box, 'x');
		box.input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', cancelable: true }));
		expect(view.state.selection).toBeInstanceOf(TextSelection);
		expect(view.state.selection.from).toBe(again + 1);
	});

	it('fills an empty equation from the gallery, or adds the built-in one after it', () => {
		const { view } = editor();
		const pos = insertNewEquation(view)!;
		placeBuiltIn(view, 'pythagorean', pos);
		expect(equations(view)).toHaveLength(1);
		expect(equations(view)[0]).toContain('<m:sSup>');
		placeBuiltIn(view, 'area', undefined);
		expect(equations(view)).toHaveLength(2);
	});

	it('refuses Linear in LaTeX input for an equation its UnicodeMath cannot hold', () => {
		const { view, box } = editor();
		const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
		const hidden = `<m:oMathPara xmlns:m="${M}"><m:oMath><m:phant><m:phantPr><m:show m:val="0"/></m:phantPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:phant></m:oMath></m:oMathPara>`;
		const pos = insertEquationXml(view, hidden, true)!;
		box.sync(view, pos, true);
		box.setFormat('latex');
		expect(box.editable).toBe(true);
		box.setLinear(true);
		expect(equations(view)[0]).toBe(hidden);
		expect(box.element.querySelector('[role="status"]')?.textContent).toContain('Linear');
	});

	it('builds up an input method composition once it ends', () => {
		const { view, box } = editor();
		box.sync(view, insertNewEquation(view)!, true);
		box.input.value = 'x';
		box.input.dispatchEvent(new InputEvent('input', { isComposing: true }));
		expect(equations(view)[0]).not.toContain('<m:t>x</m:t>');
		box.input.dispatchEvent(new CompositionEvent('compositionend'));
		expect(equations(view)[0]).toContain('<m:t>x</m:t>');
	});
});

describe('the Equation gallery over a read-only equation', () => {
	it('adds the built-in equation after it instead of replacing it', async () => {
		const toolbar = createRibbon('en');
		document.body.append(toolbar);
		const { view } = editor();
		const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
		const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const two = `<m:oMathPara xmlns:m="${M}" xmlns:w="${W}"><m:oMath><m:r><m:t>a</m:t></m:r><m:r><w:rPr><w:b/></w:rPr><m:t>b</m:t></m:r></m:oMath></m:oMathPara>`;
		insertEquationXml(view, two, true);
		const controller = new EquationController({
			view: () => view,
			toolbar: () => toolbar,
			reportError: (error) => {
				throw error;
			},
		});
		controller.sync();
		controller.handle({ type: 'equation', key: 'builtIn', value: 'area' });
		await expect.poll(() => equations(view).length).toBe(2);
		expect(equations(view)[0]).toBe(two);
	});
});

describe('the Equation tab', () => {
	it('shows while an equation is selected, with every Word group and gallery', async () => {
		const toolbar = createRibbon('en');
		document.body.append(toolbar);
		const { view } = editor('x ');
		const controller = new EquationController({
			view: () => view,
			toolbar: () => toolbar,
			reportError: (error) => {
				throw error;
			},
		});
		const panel = panelOf(toolbar, 'equation')!;
		expect(
			[...panel.querySelectorAll<HTMLElement>('.ribbon-group')].map((g) => g.dataset.label),
		).toEqual(['Tools', 'Symbols', 'Structures']);
		expect(panel.querySelectorAll('[data-equation-gallery]')).toHaveLength(13);
		const ink = panel.querySelector<HTMLButtonElement>('[aria-label="Ink Equation"]')!;
		expect(ink.disabled).toBe(true);
		expect(ink.title).toContain('handwriting recognition');
		controller.sync();
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		controller.handle({ type: 'equation', key: 'new' });
		await expect.poll(() => view.state.selection instanceof NodeSelection).toBe(true);
		controller.sync();
		expect(panel.hasAttribute('data-tab-hidden')).toBe(false);
		expect(
			panel.querySelector('[data-equation-format="unicode"]')?.getAttribute('aria-pressed'),
		).toBe('true');
		await expect.poll(() => document.querySelector('.dve-equation-editor input')).not.toBeNull();
		controller.handle({ type: 'equation', key: 'structure', value: 'matrix/m2x2' });
		await expect.poll(() => equations(view)[0] ?? '').toContain('<m:m>');
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
		controller.sync();
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		expect(document.querySelector('.dve-equation-editor')).toBeNull();
	});
});

describe('the Equation galleries', () => {
	it('draw each structure and name it in the editor language', async () => {
		const { openEquationGallery } = await import('./equation-galleries');
		const toolbar = createRibbon('fr');
		document.body.append(toolbar);
		const anchor = toolbar.querySelector<HTMLElement>('[data-equation-gallery="fraction"]')!;
		const chosen: unknown[] = [];
		openEquationGallery(anchor, 'fraction', { choose: (action) => chosen.push(action) });
		const item = document.querySelector<HTMLButtonElement>(
			'.dve-equation-gallery [aria-label="Fraction empilée"]',
		);
		expect(item?.querySelector('mfrac')).not.toBeNull();
		item!.click();
		expect(chosen).toEqual([{ type: 'equation', key: 'structure', value: 'fraction/stacked' }]);
	});
});

describe('following the selected equation', () => {
	it('stays the same equation through edits elsewhere, and not after leaving it', async () => {
		const { equationTrackPlugin, sameEquation } = await import('./equation-tab');
		const { view } = editor('ab');
		view.updateState(
			view.state.reconfigure({ plugins: [...view.state.plugins, equationTrackPlugin()] }),
		);
		insertNewEquation(view);
		expect(sameEquation(view.state)).toBe(false);
		view.dispatch(view.state.tr.insertText('zz', 1));
		expect(sameEquation(view.state)).toBe(true);
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
		expect(sameEquation(view.state)).toBe(false);
	});
});
