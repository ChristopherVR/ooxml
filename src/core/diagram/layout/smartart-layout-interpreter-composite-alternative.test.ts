import { describe, expect, it } from 'vitest';

import type { DiagramLayoutNode } from '../model';
import { isMappedSlotAlternative } from './smartart-layout-interpreter-composite-alternative';

describe('isMappedSlotAlternative', () => {
	it('false when compositeSlot is not yet resolved', () => {
		const node: DiagramLayoutNode = {
			name: 'singleCycle',
			forEachOrigin: { name: 'singleCycle' },
		};
		expect(isMappedSlotAlternative(undefined, node)).toBeFalsy();
	});

	it('false when node carries no forEachOrigin at all (a direct child)', () => {
		const compositeSlot: DiagramLayoutNode = {
			name: 'Name0',
			forEach: [{ name: 'singleCycle' }],
		};
		const node: DiagramLayoutNode = { name: 'child' };
		expect(isMappedSlotAlternative(compositeSlot, node)).toBeFalsy();
	});

	it("true when node's forEachOrigin name matches one of compositeSlot's own forEach children", () => {
		// `radial-cluster--hier5.pptx`'s exact shape: Name0's own
		// `<dgm:forEach name="singleCycle" .../>` is one of its own named slots.
		const compositeSlot: DiagramLayoutNode = {
			name: 'Name0',
			forEach: [{ name: 'textCenter' }, { name: 'singleCycle' }],
		};
		const node: DiagramLayoutNode = {
			name: 'cycle_3',
			forEachOrigin: { name: 'singleCycle', axis: ['ch'] },
		};
		expect(isMappedSlotAlternative(compositeSlot, node)).toBeTruthy();
	});

	it("false when node's forEachOrigin name matches none of compositeSlot's own forEach children", () => {
		const compositeSlot: DiagramLayoutNode = {
			name: 'Name0',
			forEach: [{ name: 'singleCycle' }],
		};
		const node: DiagramLayoutNode = {
			name: 'unrelated',
			forEachOrigin: { name: 'someOtherForEach', axis: ['ch'] },
		};
		expect(isMappedSlotAlternative(compositeSlot, node)).toBeFalsy();
	});

	it('false when compositeSlot declares no forEach children of its own', () => {
		const compositeSlot: DiagramLayoutNode = { name: 'Name0' };
		const node: DiagramLayoutNode = {
			name: 'cycle_3',
			forEachOrigin: { name: 'singleCycle', axis: ['ch'] },
		};
		expect(isMappedSlotAlternative(compositeSlot, node)).toBeFalsy();
	});
});
