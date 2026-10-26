// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { history, undo } from 'prosemirror-history';
import { EditorState, NodeSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { diagramsIn, loadDocx, type Paragraph } from 'ooxml-core/docx';
import { docToModel, modelToDoc } from './model-adapter';
import { createRibbon } from './ribbon';
import { panelOf, isTabHidden } from './ribbon-tab-api';
import { syncSmartArtDesign } from './smartart-design-ribbon';
import {
	diagramOutline,
	runDiagramEdit,
	selectedDiagram,
	setDiagramNodeText,
} from './smartart-commands';
import { SmartArtController } from './smartart-controller';

const fixture = () =>
	new Uint8Array(
		readFileSync(
			path.join(import.meta.dirname, '../../../core/docx/__fixtures__/smartart-word.docx'),
		),
	);

async function editor() {
	const loaded = await loadDocx(fixture());
	const doc = modelToDoc(loaded.model);
	let pos = -1;
	doc.descendants((node, at) => {
		if (pos < 0 && node.type.name === 'image' && node.attrs.diagram) pos = at;
	});
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({
			doc,
			selection: NodeSelection.create(doc, pos),
			plugins: [history()],
		}),
	});
	return { loaded, view, pos };
}
const texts = (view: EditorView) => selectedDiagram(view.state)!.diagram.nodes.map((n) => n.text);

describe('SmartArt editing in the Word editor', () => {
	it('applies SmartArt Design commands to the selected diagram, each one undoable', async () => {
		const { view } = await editor();
		const [alpha, beta] = selectedDiagram(view.state)!.diagram.nodes;
		const added = runDiagramEdit(view, { kind: 'add', position: 'after', nodeId: alpha!.id });
		expect(added).toBeTruthy();
		expect(texts(view)).toEqual(['Alpha', '', 'Beta', 'Gamma', 'Delta', 'Epsilon']);
		expect(selectedDiagram(view.state)!.diagram.drawing).toBeUndefined();
		runDiagramEdit(view, { kind: 'demote', nodeId: beta!.id });
		expect(diagramOutline(selectedDiagram(view.state)!.diagram).map((line) => line.level)).toEqual([
			0, 0, 1, 0, 0, 0,
		]);
		expect(runDiagramEdit(view, { kind: 'demote', nodeId: alpha!.id })).toBeUndefined();
		undo(view.state, view.dispatch);
		undo(view.state, view.dispatch);
		expect(texts(view)).toEqual(['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon']);
		expect(selectedDiagram(view.state)!.diagram.rendering).toBe('cached-drawing');
	});

	it('types node text through the model and saves it', async () => {
		const { view, loaded } = await editor();
		const beta = selectedDiagram(view.state)!.diagram.nodes[1]!;
		setDiagramNodeText(view, beta.id, 'B');
		setDiagramNodeText(view, beta.id, 'Be');
		expect(texts(view)[1]).toBe('Be');
		const model = docToModel(view.state.doc, loaded.model);
		const saved = await loaded.save(model);
		const reloaded = diagramsIn((await loadDocx(saved)).model.blocks)[0]!;
		expect(reloaded.nodes.map((node) => node.text)).toEqual([
			'Alpha',
			'Be',
			'Gamma',
			'Delta',
			'Epsilon',
		]);
	});

	it('shows the SmartArt Design tab only while a SmartArt is selected', async () => {
		const { view } = await editor();
		const ribbon = createRibbon('en');
		const panel = panelOf(ribbon, 'smartart-design')!;
		expect(isTabHidden(panel)).toBe(true);
		syncSmartArtDesign(ribbon, view.state, true, true);
		expect(isTabHidden(panel)).toBe(false);
		const toggle = panel.querySelector('.dve-smartart-text-pane-toggle')!;
		expect(toggle.getAttribute('aria-pressed')).toBe('true');
		expect(panel.querySelector<HTMLButtonElement>('[aria-label="Promote"]')!.disabled).toBe(false);
		syncSmartArtDesign(ribbon, view.state, false, false);
		expect(panel.querySelector<HTMLButtonElement>('[aria-label="Promote"]')!.disabled).toBe(true);
	});

	it('drives the text pane: Enter adds, Tab demotes, Backspace on an empty line removes', async () => {
		const { view } = await editor();
		const staged = new Map();
		const controller = new SmartArtController({
			view: () => view,
			contentWidth: () => 600,
			stage: (name, part) => staged.set(name, part),
			reportError: (error) => {
				throw error;
			},
			refresh: vi.fn(),
		});
		document.body.append(controller.pane);
		controller.setLocale('en');
		controller.handle({ type: 'smartArt', key: 'textPane' });
		expect(controller.textPaneOpen).toBe(true);
		expect(controller.pane.hidden).toBe(false);
		const inputs = () => [
			...controller.pane.shadowRoot!.querySelectorAll<HTMLInputElement>('input[data-node]'),
		];
		expect(inputs().map((input) => input.value)).toEqual([
			'Alpha',
			'Beta',
			'Gamma',
			'Delta',
			'Epsilon',
		]);
		const key = (input: HTMLInputElement, init: KeyboardEventInit) =>
			input.dispatchEvent(
				new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }),
			);
		key(inputs()[0]!, { key: 'Enter' });
		expect(texts(view)).toEqual(['Alpha', '', 'Beta', 'Gamma', 'Delta', 'Epsilon']);
		const fresh = inputs()[1]!;
		fresh.value = 'New';
		fresh.dispatchEvent(new Event('input', { bubbles: true }));
		expect(texts(view)[1]).toBe('New');
		key(inputs()[2]!, { key: 'Tab' });
		const lines = diagramOutline(selectedDiagram(view.state)!.diagram);
		expect(lines[2]).toMatchObject({ text: 'Beta', level: 1 });
		key(inputs()[2]!, { key: 'Tab', shiftKey: true });
		expect(diagramOutline(selectedDiagram(view.state)!.diagram)[2]!.level).toBe(0);
		const second = inputs()[1]!;
		second.value = '';
		second.dispatchEvent(new Event('input', { bubbles: true }));
		key(inputs()[1]!, { key: 'Backspace' });
		expect(texts(view)).toEqual(['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon']);
		controller.pane.remove();
	});

	it('runs ribbon commands on the text pane node, or the last node without one', async () => {
		const { view, loaded } = await editor();
		const controller = new SmartArtController({
			view: () => view,
			contentWidth: () => 600,
			stage: () => {},
			reportError: (error) => {
				throw error;
			},
			refresh: () => {},
		});
		controller.handle({ type: 'smartArt', key: 'addAfter' });
		expect(texts(view)).toEqual(['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon', '']);
		controller.handle({ type: 'smartArt', key: 'moveUp' });
		expect(texts(view)).toEqual(['Alpha', 'Beta', 'Gamma', 'Delta', '', 'Epsilon']);
		controller.handle({
			type: 'smartArtFormat',
			change: { kind: 'colors', value: 'monochromatic2' },
		});
		expect(selectedDiagram(view.state)!.diagram.colorsId).toBe(
			'urn:ooxml-core/colors/monochromatic2',
		);
		controller.handle({ type: 'smartArtFormat', change: { kind: 'style', value: 'moderate' } });
		expect(selectedDiagram(view.state)!.diagram.quickStyleId).toBe('urn:ooxml-core/style/moderate');
		undo(view.state, view.dispatch);
		expect(selectedDiagram(view.state)!.diagram.quickStyleId).not.toBe(
			'urn:ooxml-core/style/moderate',
		);
		const paragraph = docToModel(view.state.doc, loaded.model).blocks.find(
			(block) => block.type === 'paragraph' && block.runs.some((run) => run.image?.diagram),
		) as Paragraph;
		expect(paragraph.runs.find((run) => run.image?.diagram)?.image?.diagram?.nodes).toHaveLength(6);
	});
});
