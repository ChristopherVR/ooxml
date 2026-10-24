// @vitest-environment happy-dom
/**
 * A live colour drag in the React inspector stays on the element it started
 * on, puts back exactly what it changed, and never overwrites another edit
 * (shared `createLiveColorInput`):
 *
 * - the selection moving mid-drag puts the first shape back and the pick does
 *   not land on the second, even when both show the same colour;
 * - an edit to another field of the dragged shape survives the commit;
 * - a drag cancelled back to the start leaves an element that had no colour of
 *   its own (no line, an inherited text colour) exactly as it was.
 *
 * The real `ShapeTextPanels` over the real `useEditorHistory` and
 * `useElementOperations`, with the live edit context `ViewerInspector` provides
 * (`colour-drag.test-support.tsx`).
 */
import type { PptxElement, ShapeStyle, TextStyle } from 'pptx-viewer-core';
import { act } from 'react';
import { describe, expect, it, vi } from 'vitest';

import {
	api,
	colorInput,
	fire,
	renderHarness as render,
	setUpColourDragHarness,
} from './colour-drag.test-support';

vi.mock(import('../../i18n/use-translation'), () => ({
	useTranslation: () => ({
		t: (key: string, fallback?: unknown) => (typeof fallback === 'string' ? fallback : key),
	}),
}));

const A_STYLE: ShapeStyle = {
	fillMode: 'solid',
	fillColor: '#4472c4',
	fillColorRef: { scheme: 'accent1' },
	strokeColor: '#ed7d31',
	strokeWidth: 2,
	glowColor: '#ff0000',
	glowRadius: 8,
	glowOpacity: 0.5,
};

function shape(id: string, shapeStyle: ShapeStyle, textStyle?: TextStyle): PptxElement {
	return {
		type: 'shape',
		id,
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'rect',
		text: 'Hi',
		shapeStyle,
		...(textStyle ? { textStyle } : {}),
	} as unknown as PptxElement;
}

setUpColourDragHarness();

const fill = () => colorInput('fill');
const line = () => colorInput('line');
const text = () => colorInput('text');

describe('a fill drag the selection leaves', () => {
	it('puts the first shape back and does not land the pick on the second', () => {
		const b = {
			...A_STYLE,
			fillColor: '#00aa55',
			fillColorRef: { scheme: 'accent6' },
		} as ShapeStyle;
		render([shape('a', A_STYLE), shape('b', b)]);
		fire(fill(), 'input', '#330000');
		act(() => api!.select('b'));
		fire(fill(), 'change', '#330000');

		expect(api!.element('a').shapeStyle).toStrictEqual(A_STYLE);
		expect(api!.element('b').shapeStyle).toStrictEqual(b);
		expect(api!.canUndo()).toBeFalsy();
		expect(api!.recent()).toStrictEqual([]);
	});

	it('does not give the second shape the first one’s theme colour when both show the same hex', () => {
		const b = { ...A_STYLE, fillColorRef: { scheme: 'accent5' } } as ShapeStyle;
		render([shape('a', A_STYLE), shape('b', b)]);
		fire(fill(), 'input', '#330000');
		fire(fill(), 'input', '#4472c4');
		act(() => api!.select('b'));
		fire(fill(), 'blur', '#4472c4');

		expect(api!.element('b').shapeStyle).toStrictEqual(b);
		expect(api!.element('a').shapeStyle).toStrictEqual(A_STYLE);
	});

	it('does not write the first shape’s effects onto the second', () => {
		const b = { ...A_STYLE, glowRadius: 3, glowOpacity: 0.9 } as ShapeStyle;
		render([shape('a', A_STYLE), shape('b', b)]);
		const glow = () => colorInput('glow');
		fire(glow(), 'input', '#00ff00');
		fire(glow(), 'input', '#ff0000');
		act(() => api!.select('b'));
		fire(glow(), 'blur', '#ff0000');

		expect(api!.element('b').shapeStyle).toStrictEqual(b);
	});

	it('puts the first shape’s effect colour back on it as it is now, keeping an edit made there since', () => {
		const b = { ...A_STYLE, glowRadius: 3, glowOpacity: 0.9 } as ShapeStyle;
		render([shape('a', A_STYLE), shape('b', b)]);
		const glow = () => colorInput('glow');
		fire(glow(), 'input', '#00ff00');
		act(() => api!.edit('a', { glowRadius: 12 }));
		act(() => api!.select('b'));
		fire(glow(), 'blur', '#00ff00');

		expect(api!.element('a').shapeStyle).toStrictEqual({ ...A_STYLE, glowRadius: 12 });
		expect(api!.element('b').shapeStyle).toStrictEqual(b);
	});
});

describe('a fill drag the selection leaves after another edit to the shape', () => {
	it('puts the colour back on the shape as it is now, keeping that edit', () => {
		const b = {
			...A_STYLE,
			fillColor: '#00aa55',
			fillColorRef: { scheme: 'accent6' },
		} as ShapeStyle;
		render([shape('a', A_STYLE), shape('b', b)]);
		fire(fill(), 'input', '#330000');
		act(() => api!.edit('a', { strokeWidth: 5 }));
		act(() => api!.select('b'));
		fire(fill(), 'change', '#330000');

		expect(api!.element('a').shapeStyle).toStrictEqual({ ...A_STYLE, strokeWidth: 5 });
	});
});

describe('a text colour drag on text with no style of its own', () => {
	it('puts back only the colour, keeping what another edit added to the new style', () => {
		render([shape('a', A_STYLE)]);
		const start = text().value;
		fire(text(), 'input', '#330000');
		act(() => api!.editText('a', { bold: true }));
		fire(text(), 'input', start);
		fire(text(), 'blur', start);

		expect(api!.element('a').textStyle).toStrictEqual({ bold: true });
	});
});

describe('another edit to the dragged shape', () => {
	it('survives the commit', () => {
		render([shape('a', A_STYLE)]);
		fire(fill(), 'input', '#330000');
		act(() => api!.edit('a', { strokeWidth: 5 }));
		fire(fill(), 'change', '#330000');

		expect(api!.element('a').shapeStyle).toMatchObject({ fillColor: '#330000', strokeWidth: 5 });
	});
});

describe('a cancelled drag on an element with no colour of its own', () => {
	it('leaves a shape without a line without one', () => {
		const style: ShapeStyle = { fillMode: 'solid', fillColor: '#4472c4' };
		render([shape('a', style)]);
		const start = line().value;
		fire(line(), 'input', '#330000');
		fire(line(), 'input', start);
		fire(line(), 'blur', start);

		expect(api!.element('a').shapeStyle).toStrictEqual(style);
		expect(api!.canUndo()).toBeFalsy();
	});

	it('leaves inherited text without a colour of its own', () => {
		render([shape('a', A_STYLE, {})]);
		const start = text().value;
		fire(text(), 'input', '#330000');
		fire(text(), 'change', start);

		expect(api!.element('a').textStyle).toStrictEqual({});
		expect(api!.canUndo()).toBeFalsy();
	});
});
