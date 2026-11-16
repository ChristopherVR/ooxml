/* oxlint-disable eslint/one-var -- independent per-test locals, not intended as one statement */
import { mount } from '@vue/test-utils';
import { createCommentsPanelController } from 'ooxml-ui/pptx';
import type { PptxComment, PptxModernCommentAuthor, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';
import { defineComponent, h, ref } from 'vue';

import { COMMENTS_PANEL_CONTEXT } from '../composables/comments-panel-context';
import CommentsPanel from './CommentsPanel.vue';

const AUTHORS: PptxModernCommentAuthor[] = [{ id: '{A}', name: 'Alice', initials: 'AL' }];

function comment(overrides: Partial<PptxComment> = {}): PptxComment {
	return {
		id: 'c1',
		text: 'Looks good to me',
		author: 'Alice',
		createdAt: '2024-06-01T10:00:00Z',
		resolved: false,
		...overrides,
	};
}

/**
 * Mount the panel the way the viewer does: under the shared comments panel controller, over a
 * reactive deck whose active slide (index `active`) holds `comments`.
 */
function mountPanel(
	comments: PptxComment[],
	options: { authors?: PptxModernCommentAuthor[]; embedded?: boolean; canEdit?: boolean } = {},
) {
	const slides = ref<PptxSlide[]>([
		{ id: 's1', rId: 's1', slideNumber: 1, elements: [], comments },
		{ id: 's2', rId: 's2', slideNumber: 2, elements: [] },
	]);
	const active = ref(0);
	const controller = createCommentsPanelController({
		slides: () => slides.value,
		canEdit: () => options.canEdit ?? true,
		authorName: () => 'You',
		commit: (next) => {
			slides.value = next;
		},
	});
	const Host = defineComponent({
		setup() {
			return () =>
				h(CommentsPanel, {
					comments: slides.value[active.value]?.comments ?? [],
					authorName: 'You',
					embedded: options.embedded,
					modernCommentAuthors: options.authors ?? [],
				});
		},
	});
	const wrapper = mount(Host, {
		global: {
			provide: {
				[COMMENTS_PANEL_CONTEXT as symbol]: {
					controller,
					slideIndex: () => active.value,
					slideId: () => slides.value[active.value]?.id,
					canEdit: () => options.canEdit ?? true,
				},
			},
		},
	});
	return {
		wrapper,
		controller,
		slides: () => slides.value,
		active,
		top: () => slides.value[0]!.comments ?? [],
	};
}

const actionButton = (wrapper: ReturnType<typeof mountPanel>['wrapper'], label: string) =>
	wrapper.findAll('.pptx-comments-panel__action').find((b) => b.text() === label);

describe('commentsPanel', () => {
	it('renders an empty state when there are no comments', () => {
		const { wrapper } = mountPanel([]);
		expect(wrapper.find('[data-testid="comments-empty"]').exists()).toBeTruthy();
		expect(wrapper.find('[data-testid="comment-count"]').text()).toBe('0');
	});

	it('lists comments with author and text', () => {
		const { wrapper } = mountPanel([
			comment(),
			comment({ id: 'c2', text: 'Second', author: 'Bob' }),
		]);
		expect(wrapper.find('[data-testid="comment-count"]').text()).toBe('2');
		expect(wrapper.text()).toContain('Looks good to me');
		expect(wrapper.text()).toContain('Bob');
	});

	it('adds the trimmed draft to the slide and clears it', async () => {
		const { wrapper, top } = mountPanel([]);
		await wrapper.find('textarea').setValue('  hello world  ');
		await wrapper.find('form').trigger('submit.prevent');
		expect(top()).toStrictEqual([expect.objectContaining({ text: 'hello world', author: 'You' })]);
		expect((wrapper.find('textarea').element as HTMLTextAreaElement).value).toBe('');
	});

	it('keeps a draft per slide', async () => {
		const { wrapper, active, controller } = mountPanel([]);
		await wrapper.find('textarea').setValue('For slide one');
		active.value = 1;
		await wrapper.vm.$nextTick();
		expect((wrapper.find('textarea').element as HTMLTextAreaElement).value).toBe('');
		active.value = 0;
		await wrapper.vm.$nextTick();
		expect((wrapper.find('textarea').element as HTMLTextAreaElement).value).toBe('For slide one');
		expect(controller.getSnapshot().draftBySlideId).toStrictEqual({ s1: 'For slide one' });
	});

	it('does not add blank text and disables the button', async () => {
		const { wrapper, top } = mountPanel([]);
		const submit = wrapper.find('[data-testid="add-comment"]');
		expect(submit.attributes('disabled')).toBeDefined();
		await wrapper.find('textarea').setValue('   ');
		await wrapper.find('form').trigger('submit.prevent');
		expect(top()).toStrictEqual([]);
	});

	it('removes and resolves a comment', async () => {
		const { wrapper, top } = mountPanel([comment()]);
		await actionButton(wrapper, 'Resolve')?.trigger('click');
		expect(top()[0]!.resolved).toBeTruthy();
		expect(actionButton(wrapper, 'Reopen')).toBeDefined();
		await wrapper.find('.pptx-comments-panel__action--danger').trigger('click');
		expect(top()).toStrictEqual([]);
	});

	it('edits a comment in place', async () => {
		const { wrapper, top } = mountPanel([comment()]);
		await actionButton(wrapper, 'Edit')?.trigger('click');
		const field = wrapper.find('.pptx-comments-panel__edit');
		expect((field.element as HTMLTextAreaElement).value).toBe('Looks good to me');
		await field.setValue('Rewritten');
		await wrapper
			.findAll('button')
			.find((b) => b.text() === 'Save')
			?.trigger('click');
		expect(top()[0]!.text).toBe('Rewritten');
		expect(wrapper.find('.pptx-comments-panel__edit').exists()).toBeFalsy();
	});

	it('offers no edit on a read-only deck', () => {
		const { wrapper } = mountPanel([comment()], { canEdit: false });
		expect(actionButton(wrapper, 'Edit')).toBeUndefined();
		expect(wrapper.find('[data-testid="add-comment"]').attributes('disabled')).toBeDefined();
	});

	it('formats the timestamp via the core helper (non-empty for valid dates)', () => {
		const { wrapper } = mountPanel([comment()]);
		const time = wrapper.find('time');
		expect(time.exists()).toBeTruthy();
		expect(time.text().length).toBeGreaterThan(0);
	});

	it('threads a mention through the add-comment composer, moved with the trim', async () => {
		const { wrapper, top } = mountPanel([], { authors: AUTHORS });
		const textarea = wrapper.find('textarea').element as HTMLTextAreaElement;
		textarea.value = '  hey @al';
		textarea.setSelectionRange(9, 9);
		await wrapper.find('textarea').trigger('input');
		await wrapper.find('[data-testid="pptx-comment-mention-option"]').trigger('mousedown');
		await wrapper.find('form').trigger('submit.prevent');
		const added = top()[0]!;
		expect(added.text).toBe('hey @Alice');
		const [mention] = added.mentions ?? [];
		expect(added.text!.slice(mention!.startIndex, mention!.startIndex + mention!.length)).toBe(
			'@Alice',
		);
	});

	it('threads a mention through a reply composer', async () => {
		const { wrapper, top } = mountPanel([comment()], { authors: AUTHORS });
		await actionButton(wrapper, 'Reply')?.trigger('click');
		// The reply composer renders inside the list, before the bottom add-comment form.
		const textarea = wrapper.findAll('textarea')[0]!.element as HTMLTextAreaElement;
		textarea.value = 'ping @al';
		textarea.setSelectionRange(8, 8);
		await wrapper.findAll('textarea')[0]!.trigger('input');
		await wrapper.find('[data-testid="pptx-comment-mention-option"]').trigger('mousedown');
		await wrapper
			.findAll('button')
			.find((b) => b.text() === 'Reply' && !b.classes().includes('pptx-comments-panel__action'))
			?.trigger('click');
		const reply = top()[0]!.replies?.[0];
		expect(reply?.text).toBe('ping @Alice');
		expect(reply?.mentions?.[0]?.authorName).toBe('Alice');
	});

	it('renders its own header by default, but suppresses it when embedded', () => {
		expect(mountPanel([]).wrapper.find('.pptx-comments-panel__header').exists()).toBeTruthy();
		const embedded = mountPanel([], { embedded: true }).wrapper;
		expect(embedded.find('.pptx-comments-panel__header').exists()).toBeFalsy();
		expect(embedded.find('[data-testid="add-comment"]').exists()).toBeTruthy();
	});
});
