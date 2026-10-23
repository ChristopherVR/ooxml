import { describe, expect, it } from 'vitest';

import type { DiagramNode } from '../model/node';
import {
	addDiagramNodeAbove,
	addDiagramNodeAfter,
	addDiagramNodeBefore,
	addDiagramNodeBelow,
	removeDiagramNode,
	setDiagramNodeText,
} from './node-ops';
import { demoteDiagramNode, moveDiagramNode, promoteDiagramNode } from './node-order';
import { reflowDiagramLayout } from './reflow';
import type { DiagramEditableData, DiagramEditShape } from './types';

type Data = DiagramEditableData<DiagramNode, DiagramEditShape> & { title: string };

let counter = 0;
const options = { createModelId: () => `{N${++counter}}` };

function data(): Data {
	counter = 0;
	return {
		title: 'kept',
		nodes: [
			{ id: 'a', text: 'A', parentId: 'doc' },
			{ id: 'b', text: 'B', parentId: 'doc' },
			{ id: 'b1', text: 'B1', parentId: 'b' },
		],
		connections: [
			{ sourceId: 'doc', destId: 'a', type: 'parOf', srcOrd: 0, destOrd: 0 },
			{ sourceId: 'doc', destId: 'b', type: 'parOf', srcOrd: 1, destOrd: 0 },
			{ sourceId: 'b', destId: 'b1', type: 'parOf', srcOrd: 0, destOrd: 0 },
		],
		drawingShapes: [{ id: 's', text: 'A' }],
	};
}

const order = (next: Data) => next.nodes.map((node) => `${node.id}<${node.parentId ?? ''}`);

describe('diagram node operations', () => {
	it('adds before and after a sibling with injected ids, keeping the caller type', () => {
		const after = addDiagramNodeAfter(data(), 'New', 'a', options);
		expect(after.title).toBe('kept');
		expect(order(after)).toEqual(['a<doc', '{N1}<doc', 'b<doc', 'b1<b']);
		expect(after.connections?.at(-1)).toEqual({
			sourceId: 'doc',
			destId: '{N1}',
			type: 'parOf',
			srcOrd: 2,
			destOrd: 0,
		});
		expect(after.drawingShapes).toEqual([]);
		expect(after.drawingDirty).toBe(true);
		expect(order(addDiagramNodeBefore(data(), 'New', 'b', options))).toEqual([
			'a<doc',
			'{N1}<doc',
			'b<doc',
			'b1<b',
		]);
	});

	it('adds below (child) and above (new parent in the target place)', () => {
		expect(order(addDiagramNodeBelow(data(), 'a', 'Child', options))).toEqual([
			'a<doc',
			'b<doc',
			'b1<b',
			'{N1}<a',
		]);
		const above = addDiagramNodeAbove(data(), 'b', 'Boss', options);
		expect(order(above)).toEqual(['a<doc', '{N1}<doc', 'b<{N1}', 'b1<b']);
		expect(above.connections).toEqual([
			{ sourceId: 'doc', destId: 'a', type: 'parOf', srcOrd: 0, destOrd: 0 },
			{ sourceId: 'doc', destId: '{N1}', type: 'parOf', srcOrd: 1, destOrd: 0 },
			{ sourceId: 'b', destId: 'b1', type: 'parOf', srcOrd: 0, destOrd: 0 },
			{ sourceId: '{N1}', destId: 'b', type: 'parOf', srcOrd: 0, destOrd: 0 },
		]);
		expect(addDiagramNodeAbove(data(), 'missing', 'x', options).nodes).toHaveLength(3);
	});

	it('removes, moves, promotes and demotes', () => {
		expect(order(removeDiagramNode(data(), 'b'))).toEqual(['a<doc', 'b1<doc']);
		expect(order(moveDiagramNode(data(), 'b', -1))).toEqual(['b<doc', 'a<doc', 'b1<b']);
		expect(order(promoteDiagramNode(data(), 'b1'))).toEqual(['a<doc', 'b<doc', 'b1<doc']);
		expect(order(demoteDiagramNode(data(), 'b'))).toEqual(['a<doc', 'b<a', 'b1<b']);
	});

	it('patches the matching drawing shape text in place', () => {
		const next = setDiagramNodeText(data(), 'a', 'Alpha');
		expect(next.nodes[0]!.text).toBe('Alpha');
		expect(next.drawingShapes).toEqual([{ id: 's', text: 'Alpha' }]);
	});

	it('mints braced GUIDs by default', () => {
		const next = addDiagramNodeAfter(data(), 'x');
		expect(next.nodes.at(-1)!.id).toMatch(
			/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/,
		);
	});
});

describe('reflowDiagramLayout', () => {
	it('lays content nodes out as views in the bounds', () => {
		const shapes = reflowDiagramLayout(
			{ nodes: data().nodes, resolvedLayoutType: 'list' },
			{ x: 0, y: 0, width: 216, height: 116 },
		);
		expect(shapes?.map((shape) => [shape.id, shape.text, shape.x, shape.width])).toEqual([
			['reflow-list-a', 'A', 8, 200],
			['reflow-list-b', 'B', 8, 200],
			['reflow-list-b1', 'B1', 8, 200],
		]);
		expect(reflowDiagramLayout({ nodes: [] }, { x: 0, y: 0, width: 1, height: 1 })).toBeUndefined();
	});
});
