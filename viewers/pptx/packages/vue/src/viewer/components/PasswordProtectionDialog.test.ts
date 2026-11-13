import { mount } from '@vue/test-utils';
import { registerPptxWebControls } from 'ooxml-ui/pptx';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { ref } from 'vue';

import DialogFooter from './DialogFooter.vue';
import PasswordProtectionDialog from './PasswordProtectionDialog.vue';

vi.mock(import('vue-i18n'), () => ({
	useI18n: () => ({ t: (key: string) => key, locale: ref('en') }),
}));
beforeAll(registerPptxWebControls);

/** The dialog follows the shared password rule (`validatePasswordPair` in ooxml-ui/pptx). */
describe('PasswordProtectionDialog', () => {
	function open() {
		const wrapper = mount(PasswordProtectionDialog, {
			attachTo: document.body,
			props: { open: true, isCurrentlyProtected: false },
		});
		return wrapper;
	}

	async function submit(wrapper: ReturnType<typeof open>, password: string, confirmation: string) {
		const [first, second] = [...document.body.querySelectorAll('input')];
		first!.value = password;
		first!.dispatchEvent(new Event('input'));
		second!.value = confirmation;
		second!.dispatchEvent(new Event('input'));
		// The footer renders web controls; drive its action event, as a click on Set Password does.
		wrapper.findComponent(DialogFooter).vm.$emit('action', 'submit');
		await wrapper.vm.$nextTick();
	}

	it.each([
		['', '', 'pptx.password.errorEnter'],
		['abcd', 'abce', 'pptx.password.errorMismatch'],
		['abc', 'abc', 'pptx.password.errorTooShort'],
	])('refuses %j / %j with the shared rule', async (password, confirmation, message) => {
		const wrapper = open();
		await submit(wrapper, password, confirmation);
		expect(document.body.querySelector('[role="alert"]')?.textContent?.trim()).toBe(message);
		expect(wrapper.emitted('setPassword')).toBeUndefined();
		wrapper.unmount();
	});

	it('emits a valid password', async () => {
		const wrapper = open();
		await submit(wrapper, 'secret', 'secret');
		expect(wrapper.emitted('setPassword')).toEqual([['secret']]);
		wrapper.unmount();
	});
});
