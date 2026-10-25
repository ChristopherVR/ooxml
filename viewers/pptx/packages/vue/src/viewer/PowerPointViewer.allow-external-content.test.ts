import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { afterEach, describe, expect, it } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';

/**
 * `allowExternalContent` is a Boolean prop, so Vue would cast an omitted prop
 * to `false` and block external content the user allowed. The component must
 * keep an omitted prop `undefined`, so the Trust Center option decides.
 */
let mounted: VueWrapper | undefined;
afterEach(() => {
	mounted?.unmount();
	mounted = undefined;
	localStorage.clear();
});

describe('powerPointViewer allowExternalContent prop', () => {
	it('keeps an omitted allowExternalContent prop undefined instead of casting it to false', async () => {
		mounted = mount(PowerPointViewer, { props: { content: null } });
		await flushPromises();
		expect(mounted.props('allowExternalContent')).toBeUndefined();
	});

	it('passes explicit values through unchanged', async () => {
		mounted = mount(PowerPointViewer, { props: { content: null, allowExternalContent: false } });
		await flushPromises();
		expect(mounted.props('allowExternalContent')).toBe(false);
		await mounted.setProps({ allowExternalContent: true });
		expect(mounted.props('allowExternalContent')).toBe(true);
	});
});
