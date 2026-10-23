import { applyDiagramNodeEdit, type DiagramHostNode } from 'ooxml-core/diagram';
import { SmartArtTextPaneController } from './text-pane-controller';

afterEach(() => document.body.replaceChildren());

/** A host over a plain node list, as a product's undo history would hold it. */
function setup(initial: DiagramHostNode[] | undefined) {
	let nodes = initial;
	const toggles: boolean[] = [];
	const texts: string[] = [];
	let n = 0;
	const controller = new SmartArtTextPaneController(
		document,
		{
			nodes: () => nodes,
			edit: (edit) => {
				if (!nodes) return undefined;
				const result = applyDiagramNodeEdit(nodes, edit, { createModelId: () => `new${++n}` });
				if (!result.changed) return undefined;
				nodes = result.nodes;
				return result.nodeId ?? '';
			},
			setText: (id, text) => texts.push(`${id}=${text}`),
			toggled: (open) => toggles.push(open),
		},
		'test-pane',
	);
	document.body.append(controller.pane);
	return { controller, nodes: () => nodes, toggles, texts };
}

const detail = (pane: HTMLElement, name: string, value: unknown) =>
	pane.dispatchEvent(new CustomEvent(name, { detail: value }));

describe('SmartArtTextPaneController', () => {
	const two = () => [
		{ id: 'a', text: 'A' },
		{ id: 'b', text: 'B' },
	];

	it('opens on the selected diagram and closes from its button', () => {
		const { controller, toggles } = setup(two());
		expect(controller.pane.hidden).toBe(true);
		controller.setOpen(true);
		expect(controller.pane.hidden).toBe(false);
		expect((controller.pane as unknown as { nodes: unknown[] }).nodes).toHaveLength(2);
		detail(controller.pane, 'office-smartart-text-pane-close', {});
		expect(controller.open).toBe(false);
		expect(toggles).toEqual([true, false]);
	});

	it('stays hidden without a selected diagram', () => {
		const { controller } = setup(undefined);
		controller.setOpen(true);
		expect(controller.pane.hidden).toBe(true);
		expect(controller.runCommand('addAfter')).toBe(false);
	});

	it('turns pane requests and ribbon commands into node edits on the current node', () => {
		const { controller, nodes, texts } = setup(two());
		detail(controller.pane, 'office-smartart-node-add', { id: 'a' });
		expect(nodes()?.map((node) => node.text)).toEqual(['A', '', 'B']);
		expect(controller.currentNode()).toBe('new1');
		detail(controller.pane, 'office-smartart-node-demote', { id: 'new1' });
		expect(nodes()?.find((node) => node.id === 'new1')?.parentId).toBe('a');
		detail(controller.pane, 'office-smartart-node-select', { id: 'b' });
		expect(controller.runCommand('moveUp')).toBe(true);
		expect(nodes()?.map((node) => node.id)).toEqual(['b', 'a', 'new1']);
		expect(controller.runCommand('textPane')).toBe(false);
		detail(controller.pane, 'office-smartart-node-remove', { id: 'new1' });
		expect(nodes()?.map((node) => node.id)).toEqual(['b', 'a']);
		expect(controller.currentNode()).toBe('a');
		detail(controller.pane, 'office-smartart-text-change', { id: 'a', text: 'Z' });
		expect(texts).toEqual(['a=Z']);
	});
});
