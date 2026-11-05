import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	getStaticElementVisualState,
	getStaticElementWrapperClassName,
} from './static-element-renderer-utils';

const shape = { id: 's', type: 'shape', x: 0, y: 0, width: 10, height: 10 } as PptxElement;
const plain = { isImage: false, letsTextOverflow: false, isCallout: false, strokeWidth: 0 };

describe('getStaticElementWrapperClassName', () => {
	it('clips an unstroked shape to its box', () => {
		expect(getStaticElementWrapperClassName(shape, plain, true, false)).toContain(
			'overflow-hidden',
		);
	});

	it('does not clip a box whose text overflows it, as PowerPoint and the other bindings draw', () => {
		const overflowing = { ...plain, letsTextOverflow: true };
		expect(getStaticElementWrapperClassName(shape, overflowing, true, false)).not.toContain(
			'overflow-hidden',
		);
	});

	it('lets any box holding text overflow, not just shrink-to-fit ones', () => {
		const box = {
			...shape,
			type: 'text',
			text: 'Lorem ipsum dolor',
			textStyle: { fontSize: 58.67, textWrap: 'none' },
		} as PptxElement;
		expect(getStaticElementVisualState(box, undefined).letsTextOverflow).toBe(true);
		expect(getStaticElementVisualState(shape, undefined).letsTextOverflow).toBe(false);
	});

	it('does not clip a stroked shape, whose outline straddles the edge', () => {
		const stroked = { ...plain, strokeWidth: 4 };
		expect(getStaticElementWrapperClassName(shape, stroked, true, false)).not.toContain(
			'overflow-hidden',
		);
	});
});
