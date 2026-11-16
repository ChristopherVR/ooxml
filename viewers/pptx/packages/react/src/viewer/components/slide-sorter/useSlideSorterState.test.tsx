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

function harnessParams(overrides: Partial<Parameters<typeof useSlideSorterState>[0]> = {}) {
	return {
		slides: ['a', 'b'].map((id) => ({ id, rId: id, slideNumber: 1, elements: [] })),
		activeSlideIndex: 0,
		canEdit: true,
		sectionGroups: [],
		onSelectSlide: vi.fn(),
		onMoveSlide: vi.fn(),
		onDeleteSlides: vi.fn(),
		onDuplicateSlides: vi.fn(),
		onPasteSlides: vi.fn(),
		onToggleHideSlides: vi.fn(),
		onClose: vi.fn(),
		...overrides,
	};
}

test('keeps one keydown listener across re-renders and reaches the current onClose', () => {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const root = createRoot(target);
	const add = vi.spyOn(window, 'addEventListener');
	const remove = vi.spyOn(window, 'removeEventListener');
	function Harness({ onClose }: { onClose: () => void }): React.ReactElement {
		useSlideSorterState(harnessParams({ onClose }));
		return <div />;
	}
	const onClose = vi.fn();
	try {
		act(() => root.render(<Harness onClose={vi.fn()} />));
		act(() => root.render(<Harness onClose={onClose} />));
		expect(add.mock.calls.filter(([kind]) => kind === 'keydown')).toHaveLength(1);
		expect(remove.mock.calls.filter(([kind]) => kind === 'keydown')).toHaveLength(0);
		act(() => window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' })));
		expect(onClose).toHaveBeenCalledOnce();
	} finally {
		act(() => root.unmount());
		target.remove();
		add.mockRestore();
		remove.mockRestore();
	}
});

test('opens no menu on a read-only deck and leaves the native menu alone', () => {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const root = createRoot(target);
	let state: ReturnType<typeof useSlideSorterState>;
	function Harness(): React.ReactElement {
		state = useSlideSorterState(harnessParams({ canEdit: false }));
		return <div />;
	}
	try {
		act(() => root.render(<Harness />));
		const preventDefault = vi.fn();
		act(() =>
			state.handleContextMenu(
				{ clientX: 1, clientY: 2, preventDefault, stopPropagation: vi.fn() } as never,
				1,
			),
		);
		expect(state!.contextMenu).toBeNull();
		expect(preventDefault).not.toHaveBeenCalled();
	} finally {
		act(() => root.unmount());
		target.remove();
	}
});
