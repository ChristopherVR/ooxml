import { afterEach, describe, expect, it, vi } from 'vitest';
import { closeRibbonPopover, openMenu } from './popover';

afterEach(() => {
	closeRibbonPopover();
	document.body.innerHTML = '';
});

describe('ribbon popover focus', () => {
	it('restores a focused menu to its trigger before running the chosen command', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const root = host.attachShadow({ mode: 'open' });
		const anchor = document.createElement('button');
		root.append(anchor);
		anchor.focus();
		const run = vi.fn(() => {
			expect(root.activeElement).toBe(anchor);
		});
		const menu = openMenu(anchor, [{ label: 'Clear', run }])!;
		expect(root.activeElement).toBe(menu.querySelector('button'));
		menu.querySelector('button')!.click();
		expect(run).toHaveBeenCalledOnce();
		expect(root.activeElement).toBe(anchor);
	});
	it('leaves focus on an outside control when closing an unfocused menu', () => {
		const anchor = document.createElement('button');
		const outside = document.createElement('input');
		document.body.append(anchor, outside);
		openMenu(anchor, [{ label: 'Clear', run() {} }]);
		outside.focus();
		closeRibbonPopover();
		expect(document.activeElement).toBe(outside);
	});
});
