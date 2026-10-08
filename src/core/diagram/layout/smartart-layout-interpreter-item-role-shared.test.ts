import { describe, expect, it } from 'vitest';

import type { DiagramLayoutDefinition, DiagramLayoutNode, DiagramNode } from '../model';
import { buildConstraintIndex } from './smartart-constraint-solver';
import {
	descendantTextById,
	hasAmbiguousTopLevelRoles,
	widthWeight,
} from './smartart-layout-interpreter-item-role-shared';

describe('hasAmbiguousTopLevelRoles', () => {
	it('declines a `self` role paired with a FLAT `desOrSelf` role (Converging Radial)', () => {
		// `converging-radial--hier5.pptx`: the hub-only `centerShape` (`self`)
		// and the satellite-only `node` role (`desOrSelf`) are flattened
		// siblings of the SAME `dgm:choose`-branch-flattening shape the
		// descendant-axis check already guards - `node`'s `forEachOrigin` here
		// is `axis="self"` (trivially "the current point"), not a genuine
		// per-child repeat, so it must decline.
		const centerShape: DiagramLayoutNode = {
			name: 'centerShape',
			presentationOf: { axis: ['self'] },
			forEachOrigin: { axis: ['ch'], pointTypes: ['node'], count: [1] },
		};
		const node: DiagramLayoutNode = {
			name: 'node',
			presentationOf: { axis: ['desOrSelf'] },
			forEachOrigin: { axis: ['self'], pointTypes: ['node'] },
		};
		expect(hasAmbiguousTopLevelRoles([centerShape, node])).toBeTruthy();
	});

	it('does NOT decline a `desOrSelf` role reached through a genuine recursive `axis="ch"` forEach (Vertical Circle List)', () => {
		// `vertical-circle-list--hier5.pptx`: `txLvl3` (`desOrSelf`) is reached
		// through `axis="ch"` with NO count limit - a true per-CHILD repeat
		// (one row per indent level), not a flattened choose alternative to
		// its `self` sibling `txLvl2`. Declining here regressed the shape
		// count from 5 to 3 (both rows folded back into one).
		const txLvl2: DiagramLayoutNode = {
			name: 'txLvl2',
			presentationOf: { axis: ['self'] },
			forEachOrigin: { axis: ['ch'], pointTypes: ['node'] },
		};
		const txLvl3: DiagramLayoutNode = {
			name: 'txLvl3',
			presentationOf: { axis: ['desOrSelf'] },
			forEachOrigin: { axis: ['ch'], pointTypes: ['node'] },
		};
		expect(hasAmbiguousTopLevelRoles([txLvl2, txLvl3])).toBeFalsy();
	});

	it('declines two `des`-axis roles regardless of forEachOrigin (Continuous Cycle)', () => {
		const node1: DiagramLayoutNode = { name: 'node1', presentationOf: { axis: ['ch'] } };
		const node2: DiagramLayoutNode = { name: 'node2', presentationOf: { axis: ['des'] } };
		expect(hasAmbiguousTopLevelRoles([node1, node2])).toBeTruthy();
	});

	it('is false for a single role', () => {
		const parentText: DiagramLayoutNode = {
			name: 'parentText',
			presentationOf: { axis: ['self'] },
		};
		expect(hasAmbiguousTopLevelRoles([parentText])).toBeFalsy();
	});

	it('is false for a complementary self + non-desOrSelf descendant pair (Radial List)', () => {
		const parentNode: DiagramLayoutNode = {
			name: 'parentNode',
			presentationOf: { axis: ['self'] },
		};
		const childNode: DiagramLayoutNode = {
			name: 'childNode',
			presentationOf: { axis: ['des'] },
		};
		expect(hasAmbiguousTopLevelRoles([parentNode, childNode])).toBeFalsy();
	});
});

// Round 20: `vertical-bullet-list--hier8.pptx`'s "Branch A Root" rendered
// "Branch A Child"'s own `childText` box with "Branch A Root" DUPLICATED
// into it (COM-verified wrong) - `splitEntryFields` had no way to turn a
// `des`-axis role's resolved `nodeIds` back into real text, only
// `literalText` (transition roles) or `original.text` (the point's OWN
// text, wrong for a genuine descendant). `descendantTextById` is the fix's
// data source.
describe('descendantTextById', () => {
	it("maps every descendant's own id to its own text, not the ancestor's", () => {
		const childrenOf = new Map<string, DiagramNode[]>([
			['root', [{ id: 'child', text: 'Branch A Child', parentId: 'root' }]],
			['child', [{ id: 'grandchild', text: 'Branch A Grandchild', parentId: 'child' }]],
		]);
		const root: DiagramNode = { id: 'root', text: 'Branch A Root' };
		const map = descendantTextById(root, childrenOf);
		expect(map.get('child')).toBe('Branch A Child');
		expect(map.get('grandchild')).toBe('Branch A Grandchild');
		expect(map.has('root')).toBeFalsy(); // the point's OWN id is never its own descendant
	});

	it('is empty for a genuine leaf (no children at all)', () => {
		const childrenOf = new Map<string, DiagramNode[]>();
		const leaf: DiagramNode = { id: 'leaf', text: 'Node One' };
		expect(descendantTextById(leaf, childrenOf).size).toBe(0);
	});
});

describe('widthWeight', () => {
	it("reads the role's own arranger-declared `w` constraint (round 27's column-split mirror of `heightWeight`) - \"Vertical Bracket List\"'s own `parTx`/`desTx` shares", () => {
		const parTx: DiagramLayoutNode = { name: 'parTx' };
		const desTx: DiagramLayoutNode = { name: 'desTx' };
		const definition: DiagramLayoutDefinition = {
			rootNode: {
				name: 'linNode',
				children: [parTx, desTx],
				constraints: [
					{ type: 'w', for: 'ch', forName: 'parTx', referenceType: 'w', factor: 0.25 },
					{ type: 'w', for: 'ch', forName: 'desTx', referenceType: 'w', factor: 0.68 },
				],
			},
		};
		const index = buildConstraintIndex(definition);
		expect(widthWeight(index, 'linNode', parTx)).toBeCloseTo(0.25);
		expect(widthWeight(index, 'linNode', desTx)).toBeCloseTo(0.68);
	});

	it('defaults to an equal share (1) when no `w` constraint resolves at all', () => {
		const bare: DiagramLayoutNode = { name: 'bare' };
		const index = buildConstraintIndex({ rootNode: { name: 'root', children: [bare] } });
		expect(widthWeight(index, 'root', bare)).toBe(1);
	});
});
