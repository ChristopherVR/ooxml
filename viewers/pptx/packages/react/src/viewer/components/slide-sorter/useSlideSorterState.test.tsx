import type { PptxElement } from 'pptx-viewer-core';
// @vitest-environment happy-dom
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, vi, test } from 'vitest';

import { useSlideSorterState } from './useSlideSorterState';

test('copies a selected range and resolves its clipboard through the shared state', () => {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const root = createRoot(target);
	let state: ReturnType<typeof useSlideSorterState>;
	const duplicate = vi.fn();
	const paste = vi.fn();
	const template = {
		id: 'template',
		type: 'text',
		x: 0,
		y: 0,
		width: 100,
		height: 30,
		text: 'Captured artwork',
	} as PptxElement;
	const templates: Record<string, PptxElement[]> = { a: [template] };
	function Harness(): React.ReactElement {
		state = useSlideSorterState({
			slides: ['a', 'b', 'c'].map((id, index) => ({
				id,
				rId: id,
				slideNumber: index + 1,
				elements: [],
			})),
			activeSlideIndex: 0,
			templateElementsBySlideId: templates,
			canEdit: true,
			sectionGroups: [],
			onSelectSlide: vi.fn(),
			onMoveSlide: vi.fn(),
			onDeleteSlides: vi.fn(),
			onDuplicateSlides: duplicate,
			onPasteSlides: paste,
			onToggleHideSlides: vi.fn(),
			onClose: vi.fn(),
		});
		return <div />;
	}
	try {
		act(() => root.render(<Harness />));
		act(() => state.handleSlideClick({ shiftKey: true } as React.MouseEvent, 1));
		expect(state!.selectedIndexes).toStrictEqual([0, 1]);
		act(() => state.handleCopySelected());
		(template as PptxElement & { text: string }).text = 'Changed artwork';
		delete templates.a;
		act(() => state.handlePaste());
		expect(paste).toHaveBeenCalledWith(
			expect.objectContaining({ insertAt: 2, sourceIds: ['a', 'b'] }),
		);
		expect(duplicate).not.toHaveBeenCalled();
		expect(paste.mock.calls[0]![0].templateElements[0][0].text).toBe('Captured artwork');
		act(() => state.handleSlideClick({} as React.MouseEvent, 2));
		act(() => state.handleSlideClick({ shiftKey: true } as React.MouseEvent, 1));
		act(() => window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' })));
		act(() => state.handleSlideClick({ shiftKey: true } as React.MouseEvent, 1));
		expect(state!.selectedIndexes).toStrictEqual([0, 1]);
	} finally {
		act(() => root.unmount());
		target.remove();
	}
});
