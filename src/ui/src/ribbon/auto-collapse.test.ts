import { describe, expect, it } from 'vitest';
import { resolveCollapsed } from './auto-collapse';

describe('resolveCollapsed', () => {
	it('collapses on a short viewport and expands on a tall one without a user choice', () => {
		expect(resolveCollapsed(true, undefined)).toBe(true);
		expect(resolveCollapsed(false, undefined)).toBe(false);
	});
	it('lets an explicit choice win over the viewport', () => {
		expect(resolveCollapsed(true, false)).toBe(false);
		expect(resolveCollapsed(false, true)).toBe(true);
	});
});
