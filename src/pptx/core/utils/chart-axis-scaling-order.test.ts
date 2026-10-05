import { describe, expect, it } from 'vitest';

import type { XmlObject } from '../types';
import { applyChartAxisScaling } from './chart-axis-scaling';

const localName = (key: string) => key.replace(/^.*:/u, '');

describe('applyChartAxisScaling', () => {
	it('keeps CT_Scaling order (logBase, orientation, max, min)', () => {
		const scaling: XmlObject = { 'c:orientation': { '@_val': 'minMax' } };
		applyChartAxisScaling(scaling, { axisType: 'valAx', min: 0, max: 1, logBase: 10 }, localName);
		expect(Object.keys(scaling)).toStrictEqual(['c:logBase', 'c:orientation', 'c:max', 'c:min']);
	});

	it('leaves an authored node already in order as it is', () => {
		const max = { '@_val': '1' };
		const scaling: XmlObject = {
			'c:orientation': { '@_val': 'minMax' },
			'c:max': max,
			'c:min': { '@_val': '0' },
		};
		applyChartAxisScaling(scaling, { axisType: 'valAx', min: 0, max: 1 }, localName);
		expect(Object.keys(scaling)).toStrictEqual(['c:orientation', 'c:max', 'c:min']);
		expect(scaling['c:max']).toBe(max);
	});
});
