import { describe, expect, it } from 'vitest';

import { diagramNodeLevels } from './outline';

describe('diagramNodeLevels', () => {
	it('treats a parent outside the list (the document point) as the top level', () => {
		const nodes = [
			{ id: 'a', parentId: 'doc' },
			{ id: 'a1', parentId: 'a' },
			{ id: 'a11', parentId: 'a1' },
			{ id: 'b' },
		];
		expect(diagramNodeLevels(nodes)).toEqual([0, 1, 2, 0]);
	});
});
