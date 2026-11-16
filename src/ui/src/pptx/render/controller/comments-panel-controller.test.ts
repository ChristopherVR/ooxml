import type { PptxComment, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createCommentsPanelController, findCommentInThread } from './comments-panel-controller';

function setup(options: { canEdit?: boolean; selected?: string | null } = {}) {
	let slides: PptxSlide[] = [
		{
			id: 's1',
			rId: 's1',
			slideNumber: 1,
			elements: [],
			comments: [
				{
					id: 'c1',
					text: 'First',
					author: 'Ann',
					replies: [{ id: 'r1', text: 'Nested', author: 'Bo' } as PptxComment],
				} as PptxComment,
			],
		},
		{ id: 's2', rId: 's2', slideNumber: 2, elements: [] },
	];
	const commit = vi.fn((next: PptxSlide[]) => {
		slides = next;
	});
	const controller = createCommentsPanelController({
		slides: () => slides,
		canEdit: () => options.canEdit ?? true,
		authorName: () => 'Me',
		selectedElementId: () => options.selected ?? null,
		commit,
	});
	return {
		controller,
		commit,
		slides: () => slides,
		setSlides: (next: PptxSlide[]) => (slides = next),
	};
}

describe('comments panel controller', () => {
	it('keeps a draft per slide and adds it anchored to the selection', () => {
		const { controller, commit, slides } = setup({ selected: 'shape-1' });
		controller.setDraft('s1', '  Hello ');
		controller.setDraft('s2', 'Other slide');
		const id = controller.add(0);
		expect(id).not.toBeNull();
		expect(commit).toHaveBeenCalledTimes(1);
		const added = slides()[0]!.comments!.at(-1)!;
		expect(added).toMatchObject({ id, text: 'Hello', author: 'Me', elementId: 'shape-1' });
		expect(controller.getSnapshot().draftBySlideId).toStrictEqual({ s1: '', s2: 'Other slide' });
	});

	it('keeps recorded mentions on a plain keystroke and clears them once posted', () => {
		const { controller } = setup();
		const mention = { userId: 'u1', providerId: 'p', name: 'Ann', start: 0, length: 4 };
		controller.setDraft('s1', '@Ann', [mention as never]);
		controller.setDraft('s1', '@Ann hi');
		expect(controller.getSnapshot().draftMentionsBySlideId.s1).toHaveLength(1);
		controller.add(0);
		expect(controller.getSnapshot().draftMentionsBySlideId.s1).toBeUndefined();
	});

	it('refuses a blank draft and every command on a read-only deck', () => {
		const { controller, commit } = setup({ canEdit: false });
		controller.setDraft('s1', 'Text');
		expect(controller.add(0)).toBeNull();
		expect(controller.remove(0, 'c1')).toBe(false);
		expect(controller.toggleResolved(0, 'c1')).toBe(false);
		const blank = setup();
		blank.controller.setDraft('s1', '   ');
		expect(blank.controller.add(0)).toBeNull();
		expect(commit).not.toHaveBeenCalled();
		expect(blank.commit).not.toHaveBeenCalled();
	});

	it('edits a comment, or a nested reply, in place and refuses a blank edit', () => {
		const { controller, slides } = setup();
		controller.startEdit('s1', 'r1');
		expect(controller.getSnapshot()).toMatchObject({
			editingIdBySlideId: { s1: 'r1' },
			editDraftById: { r1: 'Nested' },
		});
		controller.setEditDraft('r1', '  ');
		expect(controller.saveEdit(0, 'r1')).toBe(false);
		controller.setEditDraft('r1', ' Rewritten ');
		expect(controller.saveEdit(0, 'r1')).toBe(true);
		expect(findCommentInThread(slides()[0]!.comments!, 'r1')!.text).toBe('Rewritten');
		expect(controller.getSnapshot().editingIdBySlideId.s1).toBeNull();
		controller.startEdit('s1', 'c1');
		controller.cancelEdit('s1');
		expect(controller.getSnapshot().editingIdBySlideId.s1).toBeNull();
		controller.startEdit('s1', 'missing');
		expect(controller.getSnapshot().editingIdBySlideId.s1).toBeNull();
	});

	it('opens one reply composer, empty, and posts the reply under its parent', () => {
		const { controller, slides } = setup();
		controller.startReply('c1');
		controller.setReplyDraft('c1', 'Answer');
		controller.startReply('c1');
		expect(controller.getSnapshot().replyDraftById.c1).toBe('');
		controller.setReplyDraft('c1', 'Answer');
		const id = controller.submitReply(0, 'c1');
		expect(id).not.toBeNull();
		expect(slides()[0]!.comments![0]!.replies!.at(-1)).toMatchObject({ id, text: 'Answer' });
		expect(controller.getSnapshot()).toMatchObject({ replyingToId: null, replyDraftById: {} });
		controller.startReply('c1');
		controller.cancelReply();
		expect(controller.getSnapshot().replyingToId).toBeNull();
	});

	it('deletes and resolves, ending an edit on that slide', () => {
		const { controller, slides } = setup();
		expect(controller.toggleResolved(0, 'c1')).toBe(true);
		expect(slides()[0]!.comments![0]!.resolved).toBe(true);
		controller.startEdit('s1', 'c1');
		expect(controller.remove(0, 'c1')).toBe(true);
		expect(slides()[0]!.comments).toStrictEqual([]);
		expect(controller.getSnapshot().editingIdBySlideId.s1).toBeNull();
	});

	it('drops the drafts and edits of a deleted slide', () => {
		const { controller, slides, setSlides } = setup();
		controller.setDraft('s1', 'Keep');
		controller.setDraft('s2', 'Gone');
		controller.startEdit('s1', 'c1');
		setSlides([slides()[0]!]);
		const before = controller.getSnapshot();
		controller.pruneDrafts();
		expect(controller.getSnapshot().draftBySlideId).toStrictEqual({ s1: 'Keep' });
		expect(controller.getSnapshot().editingIdBySlideId).toStrictEqual({ s1: 'c1' });
		const after = controller.getSnapshot();
		expect(after).not.toBe(before);
		controller.pruneDrafts();
		expect(controller.getSnapshot()).toBe(after);
	});
});
