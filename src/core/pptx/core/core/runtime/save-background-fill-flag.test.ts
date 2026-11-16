import { describe, expect, it } from 'vitest';

import type { XmlObject } from '../../types';
import { writeBackgroundFillFlag } from './save-background-fill-flag';

const sp = (attrs: XmlObject = {}): { shape: XmlObject; spPr: XmlObject } => {
	const spPr: XmlObject = { 'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } } };
	return { shape: { ...attrs, 'p:nvSpPr': {}, 'p:spPr': spPr }, spPr };
};

describe('writeBackgroundFillFlag', () => {
	it('turns the flag on and drops the fill it replaces', () => {
		const { shape, spPr } = sp();
		writeBackgroundFillFlag(shape, spPr, { useBackgroundFill: true });
		expect(shape['@_useBgFill']).toBe('1');
		expect(spPr['a:solidFill']).toBeUndefined();
	});

	it('leaves an authored flag and its spPr as loaded', () => {
		const { shape, spPr } = sp({ '@_useBgFill': '1' });
		writeBackgroundFillFlag(shape, spPr, { useBackgroundFill: true });
		expect(shape['@_useBgFill']).toBe('1');
		expect(spPr['a:solidFill']).toBeDefined();
	});

	it('removes an authored flag the user turned off', () => {
		const { shape, spPr } = sp({ '@_useBgFill': 'true' });
		writeBackgroundFillFlag(shape, spPr, { useBackgroundFill: false, fillMode: 'solid' });
		expect('@_useBgFill' in shape).toBe(false);
		expect(spPr['a:solidFill']).toBeDefined();
	});

	it('touches neither a shape without the flag nor a connector', () => {
		const { shape, spPr } = sp({ '@_useBgFill': '0' });
		writeBackgroundFillFlag(shape, spPr, {});
		expect(shape['@_useBgFill']).toBe('0');
		const connector: XmlObject = { 'p:nvCxnSpPr': {} };
		writeBackgroundFillFlag(connector, {}, { useBackgroundFill: true });
		expect(connector['@_useBgFill']).toBeUndefined();
	});
});
