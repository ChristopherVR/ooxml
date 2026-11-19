// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from 'vitest';

import { activateBackstage } from './backstage-frame';

afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

const rect = (left: number, top: number, width: number, height: number): DOMRect =>
	({
		left,
		top,
		width,
		height,
		right: left + width,
		bottom: top + height,
		x: left,
		y: top,
	}) as DOMRect;

/** A suite-like page: a sidebar, then the viewer host holding its ribbon, File tab and File page. */
function setup(): {
	sidebar: HTMLButtonElement;
	host: HTMLElement;
	fileTab: HTMLButtonElement;
	panel: HTMLElement;
	back: HTMLButtonElement;
	last: HTMLButtonElement;
} {
	const sidebar = document.createElement('button');
	const host = document.createElement('div');
	host.setAttribute('data-pptx-theme-root', '');
	const ribbon = document.createElement('div');
	const fileTab = document.createElement('button');
	const panel = document.createElement('div');
	const back = document.createElement('button');
	const last = document.createElement('button');
	panel.append(back, last);
	ribbon.append(fileTab, panel);
	host.append(ribbon);
	document.body.append(sidebar, host);
	vi.spyOn(host, 'getBoundingClientRect').mockReturnValue(rect(160, 40, 1000, 700));
	return { sidebar, host, fileTab, panel, back, last };
}

const press = (target: Element, key: string, shiftKey = false): KeyboardEvent => {
	const event = new KeyboardEvent('keydown', { key, shiftKey, bubbles: true, cancelable: true });
	target.dispatchEvent(event);
	return event;
};

describe('activateBackstage placement', () => {
	it('puts the File page on the viewer host box, not the page', () => {
		const { panel } = setup();
		const release = activateBackstage(panel, vi.fn());
		expect(panel.style.left).toBe('160px');
		expect(panel.style.top).toBe('40px');
		expect(panel.style.width).toBe('1000px');
		expect(panel.style.height).toBe('700px');
		expect(panel.style.right).toBe('auto');
		expect(panel.style.bottom).toBe('auto');
		release();
		expect(panel.getAttribute('style') ?? '').toBe('');
	});

	it('follows the host when the window resizes', () => {
		const { host, panel } = setup();
		const release = activateBackstage(panel, vi.fn());
		vi.spyOn(host, 'getBoundingClientRect').mockReturnValue(rect(200, 0, 600, 500));
		window.dispatchEvent(new Event('resize'));
		expect(panel.style.left).toBe('200px');
		expect(panel.style.width).toBe('600px');
		release();
	});

	it('corrects for a transformed ancestor that moves the fixed page', () => {
		const { panel } = setup();
		// The ancestor sits at (20, 10), so the page draws that far from where it was placed.
		vi.spyOn(panel, 'getBoundingClientRect').mockReturnValue(rect(180, 50, 1000, 700));
		const release = activateBackstage(panel, vi.fn());
		expect(panel.style.left).toBe('140px');
		expect(panel.style.top).toBe('30px');
		release();
	});

	it('leaves a page outside any viewer host alone', () => {
		const panel = document.createElement('div');
		document.body.append(panel);
		activateBackstage(panel, vi.fn())();
		expect(panel.getAttribute('style')).toBeNull();
	});
});

describe('activateBackstage keyboard', () => {
	it('closes on Escape inside the page and returns focus to the File tab', async () => {
		const { fileTab, panel, back } = setup();
		fileTab.focus();
		const onClose = vi.fn();
		const release = activateBackstage(panel, onClose);
		await Promise.resolve();
		expect(document.activeElement).toBe(back);
		expect(press(back, 'Escape').defaultPrevented).toBeTruthy();
		expect(onClose).toHaveBeenCalledOnce();
		release();
		expect(document.activeElement).toBe(fileTab);
	});

	it('leaves keys pressed outside the viewer to the page', async () => {
		const { sidebar, fileTab, panel } = setup();
		fileTab.focus();
		const onClose = vi.fn();
		const release = activateBackstage(panel, onClose);
		await Promise.resolve();
		sidebar.focus();
		expect(press(sidebar, 'Escape').defaultPrevented).toBeFalsy();
		expect(press(sidebar, 'Tab').defaultPrevented).toBeFalsy();
		expect(onClose).not.toHaveBeenCalled();
		expect(document.activeElement).toBe(sidebar);
		release();
	});

	it('keeps Tab inside the page while focus is in it', async () => {
		const { fileTab, panel, back, last } = setup();
		fileTab.focus();
		const release = activateBackstage(panel, vi.fn());
		await Promise.resolve();
		last.focus();
		press(last, 'Tab');
		expect(document.activeElement).toBe(back);
		press(back, 'Tab', true);
		expect(document.activeElement).toBe(last);
		release();
	});

	it('makes the page itself focusable so a click on empty space keeps Escape working', () => {
		const { panel } = setup();
		const release = activateBackstage(panel, vi.fn());
		expect(panel.tabIndex).toBe(-1);
		release();
	});
});
