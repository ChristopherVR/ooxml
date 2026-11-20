/**
 * Escape closes the quick-access row's drop-down menus (Presentation options, More actions) and
 * gives focus back to the button that opened them, through the shared `listenForPopupMenuDismiss`.
 */
import { mount } from '@vue/test-utils';
import { describe, expect, it } from 'vitest';
import { defineComponent, h, nextTick, ref } from 'vue';

import OverflowMenu from './OverflowMenu.vue';
import PresentDropdown from './PresentDropdown.vue';

const Overflow = defineComponent(() => {
	const open = ref(false);
	return () =>
		h(OverflowMenu, {
			isOverflowMenuOpen: open.value,
			onSetOverflowMenuOpen: (next: boolean) => {
				open.value = next;
			},
		} as never);
});

const Present = defineComponent(
	() => () => h(PresentDropdown, { isActive: false, onPresent: () => {} }),
);

describe('ribbon drop-down menus', () => {
	it.each([
		['Presentation options', Present],
		['More actions', Overflow],
	])('closes %s on Escape and gives focus back to its button', async (label, component) => {
		const wrapper = mount(component, { attachTo: document.body });
		const trigger = wrapper.get(`[aria-label="${label}"]`).element as HTMLButtonElement;
		trigger.click();
		await nextTick();
		expect(trigger.getAttribute('aria-expanded')).toBe('true');
		const escape = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
		document.body.dispatchEvent(escape);
		await nextTick();
		expect(escape.defaultPrevented).toBe(true);
		expect(trigger.getAttribute('aria-expanded')).toBe('false');
		expect(document.activeElement).toBe(trigger);
		wrapper.unmount();
	});
});
