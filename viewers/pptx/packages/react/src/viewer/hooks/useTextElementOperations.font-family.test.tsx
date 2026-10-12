// @vitest-environment happy-dom
import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';

import { useTextElementOperations } from './useTextElementOperations';

const sameFace = { fontFamily: 'Old Sans', eastAsiaFont: 'Old Sans' };

function textBox(): PptxElement {
	return {
		type: 'text',
		id: 'box',
		x: 0,
		y: 0,
		width: 100,
		height: 40,
		text: 'Sample',
		textStyle: { ...sameFace },
		textSegments: [{ text: 'Sample', style: { ...sameFace } }],
	} as PptxElement;
}

describe('useTextElementOperations font family on a whole box', () => {
	it('moves the east Asian face that matched the Latin face on every run', () => {
		const updateSelectedElement = vi.fn<(updates: Partial<PptxElement>) => void>();
		let operations: ReturnType<typeof useTextElementOperations> | undefined;
		function Harness(): null {
			operations = useTextElementOperations({
				selectedElement: textBox(),
				inlineEditingElementId: null,
				inlineEditingText: '',
				inlineEditingSnapshotRef: undefined,
				setInlineEditingElementId: () => {},
				updateSelectedElement,
			});
			return null;
		}
		const container = document.createElement('div');
		const root = createRoot(container);
		act(() => root.render(<Harness />));

		act(() => operations!.updateSelectedTextStyle({ fontFamily: 'New Serif' }));
		act(() => root.unmount());

		const patch = updateSelectedElement.mock.calls[0]?.[0] as {
			textStyle: Record<string, unknown>;
			textSegments: TextSegment[];
		};
		expect(patch.textStyle).toMatchObject({ fontFamily: 'New Serif', eastAsiaFont: 'New Serif' });
		expect(patch.textSegments[0]?.style).toMatchObject({
			fontFamily: 'New Serif',
			eastAsiaFont: 'New Serif',
		});
	});
});
