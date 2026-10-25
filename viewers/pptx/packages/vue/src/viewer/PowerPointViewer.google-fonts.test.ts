import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { afterEach, describe, expect, it } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';

/**
 * The `googleFonts` prop type includes `false`, so Vue treats it as a Boolean
 * prop and would cast an omitted prop to `false`, silently turning the Google
 * Fonts fallback off. The component must keep an omitted prop `undefined`.
 */
let mounted: VueWrapper | undefined;
afterEach(() => {
	mounted?.unmount();
	mounted = undefined;
	localStorage.clear();
});

describe('powerPointViewer googleFonts prop', () => {
	it('keeps an omitted googleFonts prop undefined instead of casting it to false', async () => {
		mounted = mount(PowerPointViewer, { props: { content: null } });
		await flushPromises();
		expect(mounted.props('googleFonts')).toBeUndefined();
	});

	it('passes explicit values through unchanged', async () => {
		mounted = mount(PowerPointViewer, { props: { content: null, googleFonts: false } });
		await flushPromises();
		expect(mounted.props('googleFonts')).toBe(false);
		await mounted.setProps({ googleFonts: { baseUrl: 'https://fonts.example.test/css2' } });
		expect(mounted.props('googleFonts')).toStrictEqual({
			baseUrl: 'https://fonts.example.test/css2',
		});
	});
});
