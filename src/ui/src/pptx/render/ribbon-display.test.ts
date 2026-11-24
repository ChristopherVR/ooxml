// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';

import {
	RIBBON_DISPLAY_HIDE_QAT_LABEL_KEY,
	RIBBON_DISPLAY_SHOW_QAT_LABEL_KEY,
	isRibbonCommandChange,
	isRibbonCommandClick,
	normalizeRibbonDisplayMode,
	ribbonDisplayLayout,
	ribbonDisplayMenu,
	ribbonDisplayQuickAccessItem,
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

	it('builds the fourth Quick Access Toolbar menu item from current visibility', () => {
		expect(ribbonDisplayQuickAccessItem(true)).toEqual({
			visible: true,
			labelKey: RIBBON_DISPLAY_HIDE_QAT_LABEL_KEY,
		});
		expect(ribbonDisplayQuickAccessItem(false)).toEqual({
			visible: false,
			labelKey: RIBBON_DISPLAY_SHOW_QAT_LABEL_KEY,
		});
	});

	it('distinguishes ribbon command clicks from tabs, dropdown triggers, disabled buttons and inputs', () => {
		const clickOn = (el: HTMLElement): boolean => {
			let result = false;
			el.addEventListener('click', (event) => {
				result = isRibbonCommandClick(event);
			});
			el.click();
			return result;
		};

		const cmd = document.createElement('button');
		expect(clickOn(cmd)).toBe(true);

		const menuItem = document.createElement('div');
		menuItem.setAttribute('role', 'menuitem');
		expect(clickOn(menuItem)).toBe(true);

		const disabled = document.createElement('button');
		disabled.disabled = true;
		let disabledResult = true;
		disabled.addEventListener('click', (e) => {
			disabledResult = isRibbonCommandClick(e);
		});
		disabled.dispatchEvent(new MouseEvent('click', { bubbles: true }));
		expect(disabledResult).toBe(false);

		const dropdown = document.createElement('button');
		dropdown.setAttribute('aria-haspopup', 'menu');
		expect(clickOn(dropdown)).toBe(false);

		const expanded = document.createElement('button');
		expanded.setAttribute('aria-expanded', 'false');
		expect(clickOn(expanded)).toBe(false);

		const tab = document.createElement('button');
		tab.setAttribute('role', 'tab');
		expect(clickOn(tab)).toBe(false);

		const menuWrap = document.createElement('div');
		menuWrap.dataset.pptxChrome = 'ribbon-display-menu';
		const menuBtn = document.createElement('button');
		menuWrap.appendChild(menuBtn);
		expect(clickOn(menuBtn)).toBe(false);

		const input = document.createElement('input');
		expect(clickOn(input)).toBe(false);

		const select = document.createElement('select');
		expect(clickOn(select)).toBe(false);
		let changeResult = false;
		select.addEventListener('change', (e) => {
			changeResult = isRibbonCommandChange(e);
		});
		select.dispatchEvent(new Event('change', { bubbles: true }));
		expect(changeResult).toBe(true);
	});
});
