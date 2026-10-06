import * as THREE from 'three';
import { describe, expect, it } from 'vitest';

import type { SmartArt3DConnector } from '../render/smartart-3d-types';
import { buildConnectorLines } from './connector-object';

const connector = (points: number, id = 'c'): SmartArt3DConnector => ({
	id,
	points: Array.from({ length: points }, (_, i) => ({ x: i * 10, y: i * 5, z: 0 })),
	color: '#336699',
	width: 1.5,
});

describe('buildConnectorLines', () => {
	it('draws one line per connector, coloured and lifted off the base plane', () => {
		const disposables: { dispose: () => void }[] = [];
		const lines = buildConnectorLines(THREE, [connector(3), connector(2, 'd')], disposables);
		expect(lines).toHaveLength(2);
		const material = lines[0]?.material as THREE.LineBasicMaterial;
		expect(material.color.getHexString()).toBe('336699');
		const positions = lines[0]?.geometry.getAttribute('position');
		expect(positions?.count).toBe(3);
		expect(positions?.getZ(0)).toBeGreaterThan(0);
		// one geometry and one material per line, all released on unmount
		expect(disposables).toHaveLength(4);
	});

	it('skips a connector that cannot form a line', () => {
		expect(buildConnectorLines(THREE, [connector(1), connector(0)], [])).toEqual([]);
	});
});
