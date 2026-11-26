import { mount } from '@vue/test-utils';
import { ribbonTabButtons } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';
import { defineComponent, h, nextTick, ref } from 'vue';

import CommentsPanel from '../CommentsPanel.vue';
import { createRibbonPropsFixture } from './ribbon-props-fixture';
import RibbonToolbar from './RibbonToolbar.vue';

/**
 * RibbonToolbar: the desktop Office-style ribbon shell. Covers the
 * `hiddenActions` gating added for issue #64: the tab bar renders the full
 * canonical `TOOLBAR_TABS` list by default and omits any tab the host lists
 * in `hiddenActions`, without disturbing the rest.
 */
describe('ribbonToolbar', () => {
	it('renders every ribbon tab by default (hiddenActions omitted)', () => {
		const wrapper = mount(RibbonToolbar, { props: createRibbonPropsFixture() });
		const tabTexts = ribbonTabButtons(wrapper.element).map((tab) => tab.textContent);
		for (const label of ['Home', 'Insert', 'Design', 'View']) {
			expect(tabTexts).toContain(label);
		}
	});

	/**
	 * The Home tab renders the Clipboard group and the Arrange group side by
	 * side, and Arrange used to repeat Cut / Copy / Paste. Two buttons on one
	 * tab answering to "Copy" is a tab that cannot be addressed by name, by a
	 * user reading it or by the cross-binding control inventory.
	 */
	it('offers each clipboard command exactly once on the Home tab', () => {
		const wrapper = mount(RibbonToolbar, {
			props: createRibbonPropsFixture({ toolbarSection: 'home' }),
		});
		const titles = wrapper.findAll('button').map((b) => b.attributes('title'));
		for (const command of ['Paste', 'Cut', 'Copy']) {
			expect(titles.filter((title) => title === command)).toHaveLength(1);
		}
	});

	it('omits a hidden tab from the tab bar', () => {
		const wrapper = mount(RibbonToolbar, {
			props: createRibbonPropsFixture({ hiddenActions: ['insert'] }),
		});
		const tabTexts = ribbonTabButtons(wrapper.element).map((tab) => tab.textContent);
		expect(tabTexts).not.toContain('Insert');
		expect(tabTexts).toContain('Home');
	});

	it('routes Drawing Arrange to the shape editor without Picture Format props', () => {
		const patch = vi.fn();
		const wrapper = mount(RibbonToolbar, {
			props: {
				...createRibbonPropsFixture({
					canEdit: true,
					selectedElement: { id: 'shape', type: 'shape', x: 0, y: 0, width: 10, height: 10 },
					selectedCount: 1,
				}),
				onPatchElement: patch,
			},
		});
		try {
			wrapper
				.get<HTMLButtonElement>('[data-ribbon-control="home.drawing.arrange"] button')
				.element.click();
			wrapper
				.get<HTMLButtonElement>('[data-ribbon-control="home.drawing.rotateLeft"]')
				.element.click();
			expect(patch).toHaveBeenCalledExactlyOnceWith('shape', { rotation: 270 });
		} finally {
			wrapper.unmount();
		}
	});

	it('keeps the header Comments button pressed while its pane is open and returns focus on close', async () => {
		const open = ref(false);
		const toolbar = createRibbonPropsFixture({
			onToggleComments: () => {
				open.value = !open.value;
			},
		});
		const Harness = defineComponent({
			setup: () => () =>
				h('div', { 'data-pptx-editor-chrome': '' }, [
					h(RibbonToolbar, { ...toolbar, isCommentsPanelOpen: open.value }),
					open.value
						? h(CommentsPanel, {
								comments: [],
								authorName: 'You',
								closeTaskPane: () => {
									open.value = false;
								},
							})
						: null,
				]),
		});
		const wrapper = mount(Harness, { attachTo: document.body });
		try {
			const actions = wrapper.get('pptx-ui-ribbon-actions').element.shadowRoot;
			const opener = actions?.querySelector<HTMLButtonElement>('[part~="comments"]');
			expect(opener?.getAttribute('aria-pressed')).toBe('false');
			opener?.focus();
			opener?.click();
			await nextTick();
			expect(opener?.getAttribute('aria-pressed')).toBe('true');
			const close = wrapper.get<HTMLButtonElement>('[data-pptx-comment-close]').element;
			close.focus();
			close.click();
			await nextTick();
			expect(opener?.getAttribute('aria-pressed')).toBe('false');
			await vi.waitFor(() => expect(actions?.activeElement).toBe(opener));
		} finally {
			wrapper.unmount();
		}
	});
});
