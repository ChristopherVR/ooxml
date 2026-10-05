import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { attachKeyTips, registerOfficeUi } from './index.js';

beforeAll(() => registerOfficeUi());
afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

function setup() {
	// Test DOMs have no layout: give every element a visible box.
	vi.spyOn(Element.prototype, 'getBoundingClientRect').mockReturnValue(
		DOMRect.fromRect({ x: 10, y: 10, width: 40, height: 20 }),
	);
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const tab = document.createElement('button');
	tab.dataset.keytip = 'H';
	tab.dataset.keytipPanel = 'home';
	const panel = document.createElement('div');
	panel.id = 'home';
	panel.dataset.keytipLevel = '';
	const bold = document.createElement('office-ui-button');
	bold.setAttribute('command', 'bold');
	bold.setAttribute('label', 'Bold');
	bold.dataset.keytip = '1';
	const paste = document.createElement('button');
	paste.dataset.keytip = 'V';
	paste.disabled = true;
	const painter = document.createElement('button');
	painter.dataset.keytip = 'FP';
	panel.append(bold, paste, painter);
	root.append(tab, panel);
	const keytips = attachKeyTips(root);
	const press = (key: string, type: 'keydown' | 'keyup' = 'keydown') =>
		tab.dispatchEvent(
			new KeyboardEvent(type, { key, bubbles: true, composed: true, cancelable: true }),
		);
	const badges = () =>
		[...root.querySelectorAll('[aria-hidden="true"] span')].map((badge) => badge.textContent);
	return { root, tab, bold, paste, painter, keytips, press, badges };
}

describe('attachKeyTips', () => {
	it('shows first-level tips on Alt, then the chosen tab panel tips', async () => {
		const { keytips, press, badges, tab } = setup();
		const clicks = vi.fn();
		tab.addEventListener('click', clicks);
		press('Alt');
		press('Alt', 'keyup');
		expect(keytips.active).toBe(true);
		expect(badges()).toEqual(['H']);
		press('h');
		expect(clicks).toHaveBeenCalledOnce();
		await new Promise((resolve) => requestAnimationFrame(resolve));
		expect(badges()).toEqual(['1', 'V', 'FP']);
	});

	it('starts and stops from code, for products that open the tips from F10', () => {
		const { keytips, badges } = setup();
		keytips.start();
		expect(keytips.active).toBe(true);
		expect(badges()).toEqual(['H']);
		keytips.stop();
		expect(keytips.active).toBe(false);
		expect(badges()).toEqual([]);
	});

	it('runs shared controls, waits for multi-letter tips and ignores disabled ones', async () => {
		const { keytips, press, bold, paste, painter } = setup();
		const seen: string[] = [];
		bold.addEventListener('office-command', (e) => seen.push((e as CustomEvent).detail.command));
		const pasted = vi.fn();
		const painted = vi.fn();
		paste.addEventListener('click', pasted);
		painter.addEventListener('click', painted);
		press('Alt');
		press('Alt', 'keyup');
		press('h');
		await new Promise((resolve) => requestAnimationFrame(resolve));
		press('v');
		expect(pasted).not.toHaveBeenCalled();
		expect(keytips.active).toBe(true);
		press('f');
		expect(painted).not.toHaveBeenCalled();
		press('p');
		expect(painted).toHaveBeenCalledOnce();
		expect(keytips.active).toBe(false);
		press('Alt');
		press('Alt', 'keyup');
		press('h');
		await new Promise((resolve) => requestAnimationFrame(resolve));
		press('1');
		expect(seen).toEqual(['bold']);
	});

	it('steps back with Escape and exits on a second Alt or a combination', () => {
		const { keytips, press, badges } = setup();
		press('Alt');
		press('Alt', 'keyup');
		press('Escape');
		expect(keytips.active).toBe(false);
		press('Alt');
		press('Tab');
		press('Alt', 'keyup');
		expect(keytips.active).toBe(false);
		press('Alt');
		press('Alt', 'keyup');
		press('Alt');
		press('Alt', 'keyup');
		expect(keytips.active).toBe(false);
		expect(badges()).toEqual([]);
		keytips.dispose();
	});
});
