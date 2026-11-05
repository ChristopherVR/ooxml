import { mount } from '@vue/test-utils';
import { describe, expect, it } from 'vitest';

import { createRibbonPropsFixture } from './ribbon-props-fixture';
import ToolbarPrimaryRow from './ToolbarPrimaryRow.vue';

describe('toolbar primary row read-only badge', () => {
	it('shows for a read-only deck and clears when editing is allowed', async () => {
		const wrapper = mount(ToolbarPrimaryRow, {
			props: createRibbonPropsFixture({ canEdit: false }),
		});
		const badge = wrapper.find('[data-pptx-read-only-badge]');
		expect(badge.exists()).toBe(true);
		expect(badge.text()).toBe('Read-only');
		expect(badge.attributes('role')).toBe('status');
		await wrapper.setProps({ canEdit: true });
		expect(wrapper.find('[data-pptx-read-only-badge]').exists()).toBe(false);
	});
});
