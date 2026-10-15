// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { history, undo } from 'prosemirror-history';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import {
	SHAPE_PALETTE,
	createDocument,
	loadDocx,
	saveDocx,
	type Paragraph,
} from 'ooxml-core/docx';
import { docToModel, modelToDoc } from './model-adapter';
import { shapeNodeView } from './shape-node-view';
import { openShapePicker } from './shape-picker';
import { ShapeTool, shapeDrawPlugin } from './shape-draw';
import { buildShapeFormatPanel, syncShapeFormat } from './shape-format-ribbon';
import { createRibbon } from './ribbon';
import { strings } from './localization-strings';
import { createShapeTextDialog } from './shape-text-dialog';
import {
	insertDrawnShape,
	selectedShape,
	setShapeFill,
	setShapeOutline,
	setShapeSize,
	setShapeText,
	setShapeWeight,
	setShapeWrap,
	shapeWrapOf,
} from './shape-commands';

function editor(text = 'Hello') {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text }] }];
	const doc = modelToDoc(model);
	const host = document.createElement('div');
	document.body.append(host);
	const tool = new ShapeTool();
	const view = new EditorView(host, {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, 2),
			plugins: [history(), shapeDrawPlugin(tool)],
		}),
		nodeViews: {
			image: (node, nodeView, getPos) =>
				node.attrs.shape
					? shapeNodeView(node, nodeView, getPos as () => number, () => undefined)
					: (undefined as never),
		},
	});
	return { model, view, tool, host };
}
const shapeRuns = (model: ReturnType<typeof docToModel>) =>
	(model.blocks[0] as Paragraph).runs.filter((run) => run.image?.shape);

describe('drawing a shape', () => {
	it('places a floating shape in front of text at the dragged rectangle and selects it', () => {
		const { view, model } = editor();
		expect(insertDrawnShape(view, 'rightArrow', 2, { x: 30, y: 12, widthPx: 140, heightPx: 70 })).toBe(
			true,
		);
		const selected = selectedShape(view.state)!;
		expect(selected).toMatchObject({ widthPx: 140, heightPx: 70, anchored: true });
		expect(selected.placement).toMatchObject({
			wrap: 'none',
			relativeFrom: 'column',
			offsetXPx: 30,
			relativeFromV: 'paragraph',
			offsetYPx: 12,
		});
		const [run] = shapeRuns(docToModel(view.state.doc, model));
		expect(run?.image).toMatchObject({ unsupported: 'Shape', anchored: true });
		expect(run?.image?.shape).toMatchObject({ preset: 'rightArrow', editable: true });
		expect((model.blocks[0] as Paragraph).runs[0]?.text).toBe('Hello');
		undo(view.state, view.dispatch);
		expect(shapeRuns(docToModel(view.state.doc, model))).toHaveLength(0);
	});

	it('uses a default size for a click and flips a line drawn up and to the left', () => {
		const { view } = editor();
		insertDrawnShape(view, 'ellipse', 2, { x: 0, y: 0 });
		expect(selectedShape(view.state)).toMatchObject({ widthPx: 96, heightPx: 64 });
		insertDrawnShape(view, 'line', 2, { x: 5, y: 5, widthPx: 80, heightPx: 40, flipH: true, flipV: true });
		expect(selectedShape(view.state)?.spec).toMatchObject({ preset: 'line', flipH: true, flipV: true });
	});

	it('arms from the gallery, marks the page for drawing and disarms', () => {
		const { view, tool } = editor();
		tool.arm('star5');
		expect(view.dom.classList.contains('dve-drawing-shape')).toBe(true);
		tool.disarm();
		expect(view.dom.classList.contains('dve-drawing-shape')).toBe(false);
	});
});

describe('Shape Format commands', () => {
	const drawn = () => {
		const ctx = editor();
		insertDrawnShape(ctx.view, 'rect', 2, { x: 10, y: 10, widthPx: 100, heightPx: 50 });
		return ctx;
	};

	it('sets fill, outline and weight', () => {
		const { view } = drawn();
		expect(setShapeFill(view, '#ff0000')).toBe(true);
		expect(selectedShape(view.state)?.spec.fill).toMatchObject({
			kind: 'solid',
			color: { value: 'FF0000' },
		});
		setShapeFill(view, 'none');
		expect(selectedShape(view.state)?.spec.fill).toEqual({ kind: 'none' });
		setShapeFill(view, 'default');
		expect(selectedShape(view.state)?.spec.fill).toBeUndefined();
		setShapeOutline(view, '#00ff00');
		setShapeWeight(view, 3);
		expect(selectedShape(view.state)?.spec.line).toMatchObject({
			widthEmu: 38100,
			fill: { kind: 'solid', color: { value: '00FF00' } },
		});
		setShapeOutline(view, 'none');
		expect(selectedShape(view.state)?.spec.line?.fill).toEqual({ kind: 'none' });
	});

	it('adds text and resizes', () => {
		const { view } = drawn();
		setShapeText(view, ['Line one', 'Line two']);
		expect(selectedShape(view.state)?.spec.text).toEqual(['Line one', 'Line two']);
		setShapeSize(view, 288, undefined);
		setShapeSize(view, undefined, 48);
		expect(selectedShape(view.state)).toMatchObject({ widthPx: 288, heightPx: 48 });
		setShapeText(view, ['']);
		expect(selectedShape(view.state)?.spec.text).toBeUndefined();
	});

	it('changes wrapping between in line, wrapped and in front of or behind text', () => {
		const { view } = drawn();
		expect(shapeWrapOf(selectedShape(view.state)!)).toBe('front');
		setShapeWrap(view, 'behind');
		expect(selectedShape(view.state)?.placement).toMatchObject({ wrap: 'none', behindText: true });
		setShapeWrap(view, 'square');
		expect(selectedShape(view.state)?.placement).toMatchObject({ wrap: 'square' });
		expect(selectedShape(view.state)?.placement?.behindText).toBeUndefined();
		setShapeWrap(view, 'inline');
		expect(selectedShape(view.state)).toMatchObject({ anchored: false });
		expect(shapeWrapOf(selectedShape(view.state)!)).toBe('inline');
		setShapeWrap(view, 'topAndBottom');
		expect(shapeWrapOf(selectedShape(view.state)!)).toBe('topAndBottom');
		expect(selectedShape(view.state)?.anchored).toBe(true);
	});

	it('refuses to edit a shape the editor cannot rewrite', () => {
		const { view } = drawn();
		const selected = selectedShape(view.state)!;
		const node = view.state.doc.nodeAt(selected.pos)!;
		view.dispatch(
			view.state.tr.setNodeMarkup(selected.pos, undefined, {
				...node.attrs,
				shape: JSON.stringify({ ...selected.spec, editable: false, textEditable: false }),
			}),
		);
		view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, selected.pos)));
		expect(setShapeFill(view, '#ff0000')).toBe(false);
		expect(setShapeText(view, ['x'])).toBe(false);
	});

	it('saves the edited shape as valid DrawingML and reads it back', async () => {
		const { view, model } = drawn();
		setShapeFill(view, '#ffc000');
		setShapeText(view, ['Saved']);
		setShapeWrap(view, 'square');
		const bytes = await saveDocx(docToModel(view.state.doc, model));
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<wp:anchor');
		expect(xml).toContain('<wp:wrapSquare wrapText="bothSides"/>');
		expect(xml).toContain('prst="rect"');
		expect(xml).toContain('val="FFC000"');
		expect(xml).toContain('<w:t>Saved</w:t>');
		const reopened = await loadDocx(bytes);
		const [run] = shapeRuns(reopened.model);
		expect(run?.image).toMatchObject({ anchored: true, widthPx: 100, heightPx: 50 });
		expect(run?.image?.shape).toMatchObject({ preset: 'rect', text: ['Saved'], editable: true });
	});
});

describe('shape node view', () => {
	it('draws the geometry as SVG at the shape size, floating or inline', () => {
		const { view, host } = editor();
		insertDrawnShape(view, 'heart', 2, { x: 20, y: 8, widthPx: 90, heightPx: 80 });
		const dom = host.querySelector<HTMLElement>('.dve-shape')!;
		expect(dom.classList.contains('dve-shape-floating')).toBe(true);
		expect(dom.style.left).toBe('20px');
		expect(dom.style.top).toBe('8px');
		expect(dom.style.width).toBe('90px');
		expect(dom.querySelector('svg path')).not.toBeNull();
		expect(dom.dataset.shapePreset).toBe('heart');
		setShapeWrap(view, 'inline');
		const inline = host.querySelector<HTMLElement>('.dve-shape')!;
		expect(inline.classList.contains('dve-shape-floating')).toBe(false);
		expect(inline.classList.contains('dve-shape-selected')).toBe(true);
	});
});

describe('shape gallery and tab', () => {
	it('lists every family and arms the chosen preset', () => {
		const anchor = document.createElement('button');
		anchor.setAttribute('aria-label', 'Shapes');
		document.body.append(anchor);
		let picked = '';
		openShapePicker(anchor, (preset) => (picked = preset), () => {});
		const pop = document.body.querySelector('.shape-picker')!;
		expect([...pop.querySelectorAll('h4')].map((h) => h.textContent)).toEqual(
			SHAPE_PALETTE.map((group) => group.name),
		);
		const buttons = [...pop.querySelectorAll<HTMLButtonElement>('button[data-preset]')];
		expect(buttons).toHaveLength(SHAPE_PALETTE.reduce((sum, group) => sum + group.shapes.length, 0));
		expect(buttons.every((button) => button.querySelector('svg path'))).toBe(true);
		buttons.find((button) => button.dataset.preset === 'wedgeRectCallout')!.click();
		expect(picked).toBe('wedgeRectCallout');
		pop.remove();
		anchor.remove();
	});

	it('shows the Shape Format tab for a selected shape and reads its values back', () => {
		const ribbon = createRibbon('en');
		document.body.append(ribbon);
		const { view } = editor();
		syncShapeFormat(ribbon, view.state, true);
		const tab = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="shape-format"]')!;
		expect(tab.hasAttribute('data-tab-hidden')).toBe(true);
		insertDrawnShape(view, 'rect', 2, { x: 0, y: 0, widthPx: 192, heightPx: 96 });
		syncShapeFormat(ribbon, view.state, true);
		expect(tab.hasAttribute('data-tab-hidden')).toBe(false);
		const input = (label: string) =>
			tab.querySelector<HTMLInputElement>(`[aria-label="${label}"], [data-localearialabel="${label}"]`)!;
		expect(input('Shape width (inches)').value).toBe('2');
		expect(input('Shape height (inches)').value).toBe('1');
		expect(input('Wrap text').value).toBe('front');
		expect(input('Shape width (inches)').disabled).toBe(false);
		syncShapeFormat(ribbon, view.state, false);
		expect(input('Shape width (inches)').disabled).toBe(true);
		ribbon.remove();
	});

	it('translates every label of the new controls', () => {
		const panel = document.createElement('div');
		buildShapeFormatPanel(panel);
		const labels = new Set<string>();
		for (const element of panel.querySelectorAll<HTMLElement>('[aria-label], [data-label], option')) {
			const label = element.getAttribute('aria-label') ?? element.dataset.label ?? element.textContent;
			if (label) labels.add(label.replace(/ controls$/, ''));
		}
		for (const group of SHAPE_PALETTE)
			for (const name of [group.name, ...group.shapes.map((shape) => shape.name)]) labels.add(name);
		for (const label of labels) expect(label in strings.en, label).toBe(true);
	});
});

describe('Add Text dialog', () => {
	it('edits the text of the selected shape', () => {
		const { view } = editor();
		insertDrawnShape(view, 'rect', 2, { x: 0, y: 0, widthPx: 100, heightPx: 50 });
		const dialog = createShapeTextDialog(() => view);
		document.body.append(dialog.element);
		dialog.open();
		const text = dialog.element.querySelector('textarea')!;
		expect(text.disabled).toBe(false);
		text.value = 'First\nSecond';
		[...dialog.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!.click();
		expect(selectedShape(view.state)?.spec.text).toEqual(['First', 'Second']);
		dialog.element.remove();
	});
});
