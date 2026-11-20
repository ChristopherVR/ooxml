import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { retainedBy } from './heap-snapshot.mjs';

const NODE_FIELDS = ['type', 'name', 'id', 'self_size', 'edge_count', 'detachedness'];
const EDGE_FIELDS = ['type', 'name_or_index', 'to_node'];
const NODE_TYPES = ['synthetic', 'object', 'native', 'string'];
const EDGE_TYPES = ['element', 'property', 'weak'];

/** Builds a snapshot from `[type, name, size, detachedness, edges: [edgeType, to][]]` rows. */
function snapshot(rows) {
	const strings = [];
	const string = (value) => {
		const at = strings.indexOf(value);
		return at >= 0 ? at : strings.push(value) - 1;
	};
	const nodes = [];
	const edges = [];
	rows.forEach(([type, name, size, detached, out], id) => {
		nodes.push(NODE_TYPES.indexOf(type), string(name), id, size, out.length, detached);
		for (const [edgeType, to] of out)
			edges.push(EDGE_TYPES.indexOf(edgeType), string('e'), to * NODE_FIELDS.length);
	});
	return {
		snapshot: {
			meta: {
				node_fields: NODE_FIELDS,
				edge_fields: EDGE_FIELDS,
				node_types: [NODE_TYPES],
				edge_types: [EDGE_TYPES],
			},
		},
		nodes,
		edges,
		strings,
	};
}

// 0 root -> 1 attached editor, 2 cache -> 3 detached editor -> 4 its state, 5 its shadow child;
// 0 -> 6 a weakly held detached editor; 2 -> 7 shared state that 3 also points at.
const heap = snapshot([
	[
		'synthetic',
		'(root)',
		0,
		0,
		[
			['element', 1],
			['element', 2],
			['weak', 6],
		],
	],
	['native', '<x-editor>', 10, 1, []],
	[
		'object',
		'Cache',
		20,
		0,
		[
			['property', 3],
			['property', 7],
		],
	],
	[
		'native',
		'<x-editor data-leak-probe="">',
		30,
		2,
		[
			['property', 4],
			['element', 5],
			['property', 7],
		],
	],
	['object', 'State', 40, 0, []],
	['native', '<div>', 50, 2, []],
	['native', '<x-editor data-leak-probe="">', 60, 2, []],
	['object', 'Shared', 70, 0, []],
]);

describe('retainedBy', () => {
	it('counts the live detached nodes whose name matches', () => {
		assert.equal(retainedBy(heap, /^<x-editor/).count, 1);
		assert.equal(retainedBy(heap, /^<span/).count, 0);
	});

	it('sums what only the selected nodes keep alive', () => {
		// The editor, its state and its child; not the shared state the cache also holds.
		assert.equal(retainedBy(heap, /^<x-editor/).retainedBytes, 30 + 40 + 50);
		assert.equal(retainedBy(heap).retainedBytes, 30 + 40 + 50);
	});

	it('refuses a snapshot that does not record detached nodes', () => {
		const old = structuredClone(heap);
		old.snapshot.meta.node_fields = NODE_FIELDS.slice(0, -1);
		assert.throws(() => retainedBy(old), /detached/);
	});
});
