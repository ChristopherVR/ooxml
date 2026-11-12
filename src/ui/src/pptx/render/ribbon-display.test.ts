import { describe, expect, it } from 'vitest';

import {
	normalizeRibbonDisplayMode,
	ribbonDisplayLayout,
	ribbonDisplayMenu,
	toggleRibbonCollapsed,
} from './ribbon-display';

describe('ribbon display options', () => {
	it('lists the modes in PowerPoint order and checks the current one', () => {
		const menu = ribbonDisplayMenu('tabs');
		expect(menu.map((item) => item.mode)).toEqual(['fullScreen', 'tabs', 'always']);
		expect(menu.filter((item) => item.checked).map((item) => item.mode)).toEqual(['tabs']);
		expect(menu.every((item) => item.labelKey.startsWith('pptx.ribbon.display.'))).toBe(true);
	});

	it('falls back to the docked ribbon for an unknown host value', () => {
		expect(normalizeRibbonDisplayMode('tabs')).toBe('tabs');
		expect(normalizeRibbonDisplayMode('compact')).toBe('always');
		expect(normalizeRibbonDisplayMode(undefined)).toBe('always');
	});

	it('toggles between the docked ribbon and the tab row, and leaves full-screen mode', () => {
		expect(toggleRibbonCollapsed('always')).toBe('tabs');
		expect(toggleRibbonCollapsed('tabs')).toBe('always');
		expect(toggleRibbonCollapsed('fullScreen')).toBe('always');
	});

	it('docks the commands, overlays them on a peek, or leaves only the reveal strip', () => {
		expect(ribbonDisplayLayout('always', false)).toEqual({
			titleBar: true,
			tabs: true,
			commands: 'docked',
			reveal: false,
		});
		expect(ribbonDisplayLayout('tabs', false).commands).toBe('hidden');
		expect(ribbonDisplayLayout('tabs', true)).toMatchObject({ tabs: true, commands: 'overlay' });
		expect(ribbonDisplayLayout('fullScreen', false)).toEqual({
			titleBar: false,
			tabs: false,
			commands: 'hidden',
			reveal: true,
		});
		expect(ribbonDisplayLayout('fullScreen', true)).toEqual({
			titleBar: false,
			tabs: true,
			commands: 'overlay',
			reveal: false,
		});
	});
});
