import { describe, expect, it } from 'vitest';

import { inspectorNumber } from './inspector-number';

describe('inspectorNumber', () => {
	it('rounds an EMU-derived px value to two decimals', () => {
		// 1.52pt = 19304 EMU, read back as px.
		expect(inspectorNumber(19304 / 9525)).toBe(2.03);
	});

	it('leaves short values as they are', () => {
		expect(String(inspectorNumber(4))).toBe('4');
		expect(String(inspectorNumber(9.6))).toBe('9.6');
	});

	it('falls back when the model has no finite value', () => {
		expect(inspectorNumber(undefined, 1)).toBe(1);
		expect(inspectorNumber(null, 14)).toBe(14);
		expect(inspectorNumber(Number.NaN, 3)).toBe(3);
		expect(inspectorNumber(undefined)).toBe(0);
	});

	it('rounds the fallback too', () => {
		expect(inspectorNumber(undefined, 1 / 3)).toBe(0.33);
	});
});
