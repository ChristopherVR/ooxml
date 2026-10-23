import { describe, it, expect, expectTypeOf } from 'vitest';

import { INSPECTOR_TABS, HEADING, CARD, INPUT, BTN } from './inspector-pane-constants';

describe('inspector-pane-constants', () => {
	describe('iNSPECTOR_TABS', () => {
		it('is a non-empty array', () => {
			expect(Array.isArray(INSPECTOR_TABS)).toBeTruthy();
			expect(INSPECTOR_TABS.length).toBeGreaterThan(0);
		});

		it('has no duplicate keys', () => {
			const keys = INSPECTOR_TABS.map((t) => t.key);
			expect(new Set(keys).size).toBe(keys.length);
		});

		it('contains expected tab keys', () => {
			const keys = INSPECTOR_TABS.map((t) => t.key);
			expect(keys).toContain('elements');
			expect(keys).toContain('properties');
			expect(keys).toContain('comments');
		});

		it('every entry has key, label key, and icon', () => {
			for (const tab of INSPECTOR_TABS) {
				expectTypeOf(tab.key).toBeString();
				expect(tab.key.length).toBeGreaterThan(0);
				expectTypeOf(tab.labelKey).toBeString();
				expect(tab.labelKey.length).toBeGreaterThan(0);
				expect(tab.icon).toBeDefined();
			}
		});

		it('has no duplicate label keys', () => {
			const labels = INSPECTOR_TABS.map((t) => t.labelKey);
			expect(new Set(labels).size).toBe(labels.length);
		});
	});

	describe('cSS class tokens', () => {
		it('hEADING is a non-empty string', () => {
			expectTypeOf(HEADING).toBeString();
			expect(HEADING.length).toBeGreaterThan(0);
		});

		it('cARD is a non-empty string', () => {
			expectTypeOf(CARD).toBeString();
			expect(CARD.length).toBeGreaterThan(0);
		});

		it('iNPUT is a non-empty string', () => {
			expectTypeOf(INPUT).toBeString();
			expect(INPUT.length).toBeGreaterThan(0);
		});

		it('bTN is a non-empty string', () => {
			expectTypeOf(BTN).toBeString();
			expect(BTN.length).toBeGreaterThan(0);
		});
	});
});
