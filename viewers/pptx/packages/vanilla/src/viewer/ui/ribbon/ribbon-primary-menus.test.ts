import { afterEach, describe, expect, it, vi } from 'vitest';

import { makePrimaryMenu } from './ribbon-primary-menus';

afterEach(() => {
	document.body.replaceChildren();
});

function mountMenu() {
	const host = document.createElement('div');
	const trigger = document.createElement('button');
	trigger.setAttribute('aria-haspopup', 'menu');
	const run = vi.fn();
	const menu = makePrimaryMenu(document, 'Presentation options', [{ label: 'Present', run }]);
	host.append(trigger, menu.el);
	document.body.append(host);
	return { trigger, menu, run };
}

const escape = (): KeyboardEvent =>
	new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });

describe('makePrimaryMenu', () => {
	it('closes on Escape from inside the menu and gives focus back to its button', () => {
		const { trigger, menu } = mountMenu();
		menu.toggle();
		const item = menu.el.querySelector<HTMLButtonElement>('[role="menuitem"]')!;
		item.focus();
		const event = escape();
		item.dispatchEvent(event);
		expect(event.defaultPrevented).toBe(true);
		expect(menu.isOpen()).toBe(false);
		expect(menu.el.hidden).toBe(true);
		expect(document.activeElement).toBe(trigger);
	});

	it('leaves Escape alone once the menu is closed', () => {
		const { menu } = mountMenu();
		menu.toggle();
		menu.close();
		const event = escape();
		document.body.dispatchEvent(event);
		expect(event.defaultPrevented).toBe(false);
	});
});
