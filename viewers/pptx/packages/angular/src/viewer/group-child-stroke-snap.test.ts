import { readFileSync } from 'node:fs';
import path from 'node:path';

import { getGroupChildParentFill } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { buildConnectorGeometry } from './connector-path';

/**
 * A group child's strokes snap to the device grid by their ABSOLUTE position
 * (group origin plus the child's offset), as PowerPoint's export does. There is
 * no TestBed here (see `vitest.config.ts`): the group renderer is read from the
 * source and the connector geometry is exercised directly.
 */
describe('angular: group child strokes snap by absolute position', () => {
	it('links a group to its children where it renders them', () => {
		const source = readFileSync(path.join(__dirname, 'element-renderer.component.ts'), 'utf8');
		expect(source).toMatch(
			/getGroupChildParentFill\(this\.element\(\), this\.parentGroupFill\(\)\)/u,
		);
	});

	it('snaps a connector inside a group at group y + child y', () => {
		const child = {
			type: 'connector',
			id: 'c1',
			x: 10,
			y: 23.4,
			width: 50,
			height: 0,
			shapeStyle: { strokeColor: '#000000', strokeWidth: 0.3333 },
		} as unknown as PptxElement;
		const group = {
			type: 'group',
			id: 'g1',
			x: 200,
			y: 100,
			width: 300,
			height: 100,
			children: [child],
		} as unknown as PptxElement;
		getGroupChildParentFill(group);
		const snap = buildConnectorGeometry(child, 0).snapTransform;
		expect(snap).toContain('round(123.4px');
		expect(snap).not.toContain('round(23.4px');
	});
});
