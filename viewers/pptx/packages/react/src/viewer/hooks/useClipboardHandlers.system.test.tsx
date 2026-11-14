// @vitest-environment happy-dom
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { ClipboardHandlers } from './element-manipulation-types';
import { useClipboardHandlers } from './useClipboardHandlers';

const box = (id: string, text: string) =>
	({ id, type: 'text', x: 0, y: 0, width: 10, height: 10, text }) as unknown as PptxElement;

afterEach(() => {
	vi.unstubAllGlobals();
	document.body.innerHTML = '';
});

describe('React copy writes the system clipboard', () => {
	it('keeps the internal payload and offers the selection text to other apps', async () => {
		vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
		const writeText = vi.fn(async () => undefined);
		vi.stubGlobal('ClipboardItem', undefined);
		vi.stubGlobal('navigator', { clipboard: { writeText } });
		const a = box('a', 'Alpha');
		const slide = { id: 's1', elements: [a, box('b', 'Beta')] } as unknown as PptxSlide;
		const setClipboardPayload = vi.fn();
		let handlers: ClipboardHandlers | undefined;
		function Probe() {
			handlers = useClipboardHandlers({
				activeSlide: slide,
				activeSlideIndex: 0,
				selectedElement: a,
				effectiveSelectedIds: ['a', 'b'],
				editTemplateMode: false,
				clipboardPayload: null,
				setClipboardPayload,
				ops: {} as never,
				history: {} as never,
			});
			return null;
		}
		const container = document.createElement('div');
		document.body.append(container);
		const root = createRoot(container);
		act(() => root.render(<Probe />));
		handlers?.handleCopy();
		expect(setClipboardPayload).toHaveBeenCalledOnce();
		await vi.waitFor(() => expect(writeText).toHaveBeenCalledWith('Alpha\nBeta'));
		act(() => root.unmount());
	});
});
