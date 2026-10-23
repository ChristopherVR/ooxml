import { describe, expect, it } from 'vitest';
import { parseDiagramDataModel } from '../data-model';
import { writeInsertedDiagramData } from '../write/inserted';
import { DIAGRAM_DESIGN_COMMANDS, diagramDesignEdit } from './design-commands';
import { applyDiagramNodeEdit, createDiagramNodes, diagramNodeOutline } from './node-edit';
import { diagramNodesInTreeOrder } from './node-tree';

let n = 0;
const createModelId = () => `{00000000-0000-4000-8000-${String(++n).padStart(12, '0')}}`;

describe('applyDiagramNodeEdit', () => {
	const nodes = [
		{ id: 'a', text: 'A' },
		{ id: 'b', text: 'B' },
	];

	it('adds after, keeps tree order and reports the new node', () => {
		const result = applyDiagramNodeEdit(
			nodes,
			{ kind: 'add', position: 'after', nodeId: 'a' },
			{
				createModelId,
			},
		);
		expect(result.changed).toBe(true);
		expect(result.nodes.map((node) => node.text)).toEqual(['A', '', 'B']);
		expect(result.nodeId).toBe(result.nodes[1]?.id);
	});

	it('demotes under the previous sibling and promotes back', () => {
		const demoted = applyDiagramNodeEdit(nodes, { kind: 'demote', nodeId: 'b' });
		expect(demoted.nodes.find((node) => node.id === 'b')?.parentId).toBe('a');
		const promoted = applyDiagramNodeEdit(demoted.nodes, { kind: 'promote', nodeId: 'b' });
		expect(promoted.nodes.find((node) => node.id === 'b')?.parentId).toBeUndefined();
	});

	it('reports a command that does not apply as unchanged', () => {
		const result = applyDiagramNodeEdit(nodes, { kind: 'demote', nodeId: 'a' });
		expect(result).toMatchObject({ changed: false, nodeId: 'a' });
	});

	it('drops rich paragraphs only when the text changes', () => {
		const rich = [
			{ id: 'a', text: 'A', paragraphs: [] },
			{ id: 'b', text: 'B', paragraphs: [] },
		];
		const result = applyDiagramNodeEdit(rich, { kind: 'text', nodeId: 'a', text: 'Z' });
		expect(result.nodes[0]).toEqual({ id: 'a', text: 'Z' });
		expect(result.nodes[1]).toEqual({ id: 'b', text: 'B', paragraphs: [] });
	});
});

describe('createDiagramNodes and the outline', () => {
	it('puts a hierarchy under one top node', () => {
		const created = createDiagramNodes(['Top', 'L', 'R'], 'hierarchy', { createModelId });
		expect(diagramNodeOutline(created).map((line) => line.level)).toEqual([0, 1, 1]);
		expect(diagramNodeOutline(createDiagramNodes(['1', '2'], 'list')).map((l) => l.level)).toEqual([
			0, 0,
		]);
	});
});

describe('design commands', () => {
	it('maps every node command to an edit and leaves the others alone', () => {
		for (const command of DIAGRAM_DESIGN_COMMANDS) {
			const edit = diagramDesignEdit(command.key, 'x');
			if (command.key === 'textPane' || command.key === 'changeLayout')
				expect(edit).toBeUndefined();
			else expect(edit).toBeDefined();
		}
		expect(diagramDesignEdit('moveUp', 'x')).toEqual({ kind: 'move', nodeId: 'x', direction: -1 });
		expect(diagramDesignEdit('addBelow', undefined)).toEqual({ kind: 'add', position: 'below' });
	});
});

describe('writeInsertedDiagramData', () => {
	it('writes the nodes and definitions, and reads back in the same tree order', () => {
		const nodes = createDiagramNodes(['Top', 'Left', 'Right'], 'hierarchy', { createModelId });
		const xml = writeInsertedDiagramData(
			{
				nodes,
				layoutId: 'urn:microsoft.com/office/officeart/2005/8/layout/orgChart1',
				layoutCategory: 'hierarchy',
				colorsId: 'urn:microsoft.com/office/officeart/2005/8/colors/accent1_2',
				quickStyleId: 'urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1',
			},
			createModelId,
		);
		expect(xml).toContain('loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/orgChart1"');
		expect(xml).not.toContain('type="pres"');
		const read = diagramNodesInTreeOrder(parseDiagramDataModel(xml));
		expect(diagramNodeOutline(read).map((line) => [line.text, line.level])).toEqual([
			['Top', 0],
			['Left', 1],
			['Right', 1],
		]);
	});
});
