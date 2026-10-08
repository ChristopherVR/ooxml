import { describe, expect, it } from 'vitest';

import type { DiagramLayoutNode } from '../model';
import { resolveNonHierarchyChoose } from './smartart-layout-interpreter-choose-resolve';
import type { RawXmlView } from './smartart-choose-xml';

/** These nodes carry no raw XML; the walkers see none, as before the move. */
const noXml: RawXmlView<unknown> = () => undefined;

/**
 * `radial-cluster--hier5.pptx`'s exact shape: `cycle_3` is reached through
 * `Name0`'s own `<dgm:forEach name="singleCycle" .../>` and resolves a
 * `cycle` algorithm through its OWN `dgm:choose`. Once `Name0` has already
 * resolved as the diagram's genuine top-level `compositeSlot`, `cycle_3`'s
 * own resolved algorithm must NOT win `chosen` - it describes that ONE
 * slot's own small internal arrangement (ROUND 42's `isMappedSlotAlternative`
 * exclusion), not a competing whole-diagram algorithm.
 */
describe('resolveNonHierarchyChoose (round 42: composite slot alternative exclusion)', () => {
	function slotChild(): DiagramLayoutNode {
		return { name: 'cycle_3', forEachOrigin: { name: 'singleCycle', axis: ['ch'] } };
	}

	it('wins chosen when no compositeSlot has been resolved yet', () => {
		const resolution = resolveNonHierarchyChoose(
			noXml,
			slotChild(),
			{ type: 'cycle' },
			5,
			{},
			new Set(),
			undefined,
		);
		expect(resolution.chosen?.kind).toBe('cycle');
	});

	it("does NOT win chosen once compositeSlot is resolved and node's forEachOrigin matches one of its own named forEach children", () => {
		const compositeSlot: DiagramLayoutNode = {
			name: 'Name0',
			algorithm: { type: 'composite' },
			forEach: [{ name: 'textCenter' }, { name: 'singleCycle' }],
		};
		const resolution = resolveNonHierarchyChoose(
			noXml,
			slotChild(),
			{ type: 'cycle' },
			5,
			{},
			new Set(),
			compositeSlot,
		);
		expect(resolution.chosen).toBeUndefined();
	});

	it('still wins chosen when compositeSlot is resolved but the node is an UNRELATED sibling (no matching forEach name)', () => {
		const compositeSlot: DiagramLayoutNode = {
			name: 'Name0',
			algorithm: { type: 'composite' },
			forEach: [{ name: 'someOtherSlot' }],
		};
		const resolution = resolveNonHierarchyChoose(
			noXml,
			slotChild(),
			{ type: 'cycle' },
			5,
			{},
			new Set(),
			compositeSlot,
		);
		expect(resolution.chosen?.kind).toBe('cycle');
	});
});
