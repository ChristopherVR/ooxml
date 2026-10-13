// @vitest-environment happy-dom
import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';

import { useTextElementOperations } from './useTextElementOperations';

function authoredElement(): PptxElement {
	return {
		type: 'text',
		id: 't1',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		text: 'hello',
		textStyle: {},
		textSegments: [
			{
				text: 'hello',
				style: {},
				paragraphProperties: { lineSpacingExactPt: 12, paragraphMarginLeft: 0 },
			},
		],
		paragraphIndents: [{ marginLeft: 0 }],
	} as PptxElement;
}

describe('useTextElementOperations paragraph formatting', () => {
	it('reaches line spacing and indent a paragraph authored on its own a:pPr', () => {
		const updateSelectedElement = vi.fn();
		let ops: ReturnType<typeof useTextElementOperations> | undefined;
		function Probe(): null {
			ops = useTextElementOperations({
				selectedElement: authoredElement(),
				inlineEditingElementId: null,
				inlineEditingText: '',
				setInlineEditingElementId: vi.fn(),
				updateSelectedElement,
			});
			return null;
		}
		globalThis.IS_REACT_ACT_ENVIRONMENT = true;
		const container = document.createElement('div');
		const root = createRoot(container);
		try {
			act(() => root.render(<Probe />));
			ops?.updateSelectedTextStyle({ lineSpacing: 2, paragraphMarginLeft: 24 });
		} finally {
			act(() => root.unmount());
			globalThis.IS_REACT_ACT_ENVIRONMENT = false;
		}
		const patch = updateSelectedElement.mock.calls[0]?.[0] as {
			textSegments?: TextSegment[];
			paragraphIndents?: unknown;
		};
		expect(patch.textSegments?.[0]?.paragraphProperties).toStrictEqual({
			paragraphMarginLeft: 24,
			lineSpacing: 2,
		});
		expect(patch.paragraphIndents).toStrictEqual([{ marginLeft: 24 }]);
	});
});
