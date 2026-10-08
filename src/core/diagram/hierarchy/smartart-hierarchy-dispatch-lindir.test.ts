import { describe, expect, it } from 'vitest';

import type { DiagramLayoutNode } from '../model';
import type { RawXmlView } from '../layout/smartart-choose-xml';
import {
	resolveHierarchyDispatchChAlign,
	resolveHierarchyDispatchLinDir,
	resolveHierarchyRootAlign,
} from './smartart-hierarchy-dispatch-lindir';

/** These layout-node literals carry no raw `dgm:choose` XML. */
const view: RawXmlView<unknown> = () => undefined;

/** `hierarchy-list--hier5.pptx`'s own real shape: an outer `hierChild` (mirror
 * only, `linDir="fromL"/"fromR"`), a nested `hierRoot` with `hierAlign`, and
 * THAT node's own nested `hierChild` (the genuine "children of root" generation).
 * Uses a direct `.algorithm` (not `dgm:choose`) - the choose-resolution path
 * itself (`chooseAlgorithm`) has its own dedicated test coverage elsewhere;
 * this module's own job is the STRUCTURAL search below, exercised either way. */
function cornerTree(): DiagramLayoutNode {
	return {
		name: 'diagram',
		algorithm: { type: 'hierChild', parameters: [{ type: 'linDir', value: 'fromL' }] },
		children: [
			{
				name: 'root',
				algorithm: { type: 'hierRoot', parameters: [{ type: 'hierAlign', value: 'tL' }] },
				children: [
					{
						name: 'childShape',
						algorithm: {
							type: 'hierChild',
							parameters: [
								{ type: 'chAlign', value: 'l' },
								{ type: 'linDir', value: 'fromT' },
							],
						},
						children: [],
					},
				],
			},
		],
	};
}

/** The plain "Hierarchy" family's own shape: `hierRoot1` nests NO `hierChild` at all -
 * the next generation's `hierChildN` is a SIBLING, not a descendant. */
function plainFanningTree(): DiagramLayoutNode {
	return {
		name: 'hierChild1',
		algorithm: { type: 'hierChild', parameters: [{ type: 'linDir', value: 'fromL' }] },
		children: [
			{
				name: 'hierRoot1',
				algorithm: { type: 'hierRoot' },
				children: [{ name: 'composite', algorithm: { type: 'composite' }, children: [] }],
			},
			{
				name: 'hierChild2',
				algorithm: { type: 'hierChild', parameters: [{ type: 'linDir', value: 'fromL' }] },
				children: [],
			},
		],
	};
}

describe('resolveHierarchyDispatchLinDir', () => {
	it('resolves the nested hierChild (found below the first hierRoot), not the outermost choose-wrapped one', () => {
		expect(resolveHierarchyDispatchLinDir(view, cornerTree(), 1, undefined)).toBe('fromT');
	});

	it('returns undefined for the plain fanning family (hierRoot nests no hierChild)', () => {
		expect(resolveHierarchyDispatchLinDir(view, plainFanningTree(), 1, undefined)).toBeUndefined();
	});

	it('returns undefined for an undefined algorithmNode', () => {
		expect(resolveHierarchyDispatchLinDir(view, undefined, 1, undefined)).toBeUndefined();
	});
});

describe('resolveHierarchyDispatchChAlign', () => {
	it("resolves the SAME nested hierChild's own chAlign param", () => {
		expect(resolveHierarchyDispatchChAlign(view, cornerTree(), 1, undefined)).toBe('l');
	});
});

describe('resolveHierarchyRootAlign', () => {
	it("resolves the tree's own hierRoot hierAlign param, even when reached only via a nested search", () => {
		expect(resolveHierarchyRootAlign(view, cornerTree(), 1, undefined)).toBe('tL');
	});

	it('returns undefined when no hierRoot declares hierAlign', () => {
		const tree: DiagramLayoutNode = {
			name: 'hierRoot1',
			algorithm: { type: 'hierRoot' },
			children: [],
		};
		expect(resolveHierarchyRootAlign(view, tree, 1, undefined)).toBeUndefined();
	});
});
