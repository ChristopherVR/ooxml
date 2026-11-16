// @vitest-environment happy-dom
import type { PptxSlide } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, test, vi } from 'vitest';

import { useComments } from './useComments';

test('keeps drafts per slide on the shared controller and commits through the viewer', () => {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const root = createRoot(target);
	let slides: PptxSlide[] = [
		{ id: 'a', rId: 'a', slideNumber: 1, elements: [] },
		{ id: 'b', rId: 'b', slideNumber: 2, elements: [] },
	];
	const markDirty = vi.fn();
	let result: ReturnType<typeof useComments>;
	function Harness(): React.ReactElement {
		result = useComments({
			slides,
			activeSlideIndex: 0,
			canEdit: true,
			userName: 'Me',
			selectedElementId: 'shape',
			onUpdateSlides: (updater) => {
				slides = updater(slides);
				render();
			},
			onMarkDirty: markDirty,
		});
		return <div />;
	}
	const render = (): void => act(() => root.render(<Harness />));
	try {
		render();
		act(() => result.handleCommentDraftChange('a', 'On A'));
		act(() => result.handleCommentDraftChange('b', 'On B'));
		act(() => result.handleAddSlideComment(0));
		expect(slides[0]!.comments).toStrictEqual([
			expect.objectContaining({ text: 'On A', author: 'Me', elementId: 'shape' }),
		]);
		expect(markDirty).toHaveBeenCalledOnce();
		expect(result!.commentDraftBySlideId).toStrictEqual({ a: '', b: 'On B' });
		const id = slides[0]!.comments![0]!.id;
		act(() => result.handleStartCommentEdit('a', id));
		expect(result!.commentEditDraftByCommentId[id]).toBe('On A');
		// Deleting slide b drops its draft.
		slides = [slides[0]!];
		render();
		expect(result!.commentDraftBySlideId).toStrictEqual({ a: '' });
	} finally {
		act(() => root.unmount());
		target.remove();
	}
});
