import { describe, expect, it } from 'vitest';

import type { DiagramLayoutNode, DiagramNode } from '../model';
import { compoundChildIds } from './smartart-layout-interpreter-item-role-compound';

function childrenOf(entries: Record<string, DiagramNode[]>): Map<string, DiagramNode[]> {
	return new Map(Object.entries(entries));
}

describe('compoundChildIds', () => {
	const one: DiagramNode = { id: 'one', text: 'Node One' };
	const four: DiagramNode = { id: 'four', text: 'Node Four' };
	const five: DiagramNode = { id: 'five', text: 'Node Five' };
	const map = childrenOf({ one: [four, five], four: [five] });

	it('returns [] when the axis has no "ch" token at all (nothing to position)', () => {
		const role: DiagramLayoutNode = { presentationOf: { axis: ['des', 'self'] } };
		expect(compoundChildIds(role, one, childrenOf({ one: [four] }))).toStrictEqual([]);
	});

	it("folds a selected position's own descendants in alongside it (desOrSelf second token)", () => {
		// four's own child (five) folds into four's entry when four is the
		// selected position - table-list--hier5.pptx's pillarX pattern.
		const role: DiagramLayoutNode = {
			presentationOf: { axis: ['ch', 'desOrSelf'], start: [1, 1], count: [1, 0] },
		};
		expect(compoundChildIds(role, one, map)).toStrictEqual(['four', 'five']);
	});

	it('skips a resolved position whose own text is empty', () => {
		const blank: DiagramNode = { id: 'blank', text: '' };
		const role: DiagramLayoutNode = {
			presentationOf: { axis: ['ch', 'desOrSelf'], start: [1, 1], count: [1, 0] },
		};
		expect(compoundChildIds(role, one, childrenOf({ one: [blank] }))).toStrictEqual([]);
	});
});
