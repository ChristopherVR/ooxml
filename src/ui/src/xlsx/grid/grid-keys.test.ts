import { describe, expect, it } from 'vitest';
import { navigationKey } from './grid-keys';

const key = (name: string, ctrlKey = false, metaKey = false) => ({ key: name, ctrlKey, metaKey });

describe('navigationKey', () => {
	it('moves by a screen on PageUp and PageDown', () => {
		expect(navigationKey(key('PageUp'))).toBe('pageUp');
		expect(navigationKey(key('PageDown'))).toBe('pageDown');
	});

	it('leaves Ctrl+PageUp and Ctrl+PageDown to sheet switching', () => {
		expect(navigationKey(key('PageUp', true))).toBeUndefined();
		expect(navigationKey(key('PageDown', true))).toBeUndefined();
		expect(navigationKey(key('PageDown', false, true))).toBeUndefined();
	});
});
