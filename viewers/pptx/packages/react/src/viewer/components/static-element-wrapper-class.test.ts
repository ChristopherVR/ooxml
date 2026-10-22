import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { getStaticElementWrapperClassName } from './static-element-renderer-utils';

const shape = { id: 's', type: 'shape', x: 0, y: 0, width: 10, height: 10 } as PptxElement;
const plain = { isImage: false, letsTextOverflow: false, isCallout: false, strokeWidth: 0 };

describe('getStaticElementWrapperClassName', () => {
	it('clips an unstroked shape to its box', () => {
		expect(getStaticElementWrapperClassName(shape, plain, true, false)).toContain(
			'overflow-hidden',
		);
	});

	it('does not clip a stroked shape, whose outline straddles the edge', () => {
		const stroked = { ...plain, strokeWidth: 4 };
		expect(getStaticElementWrapperClassName(shape, stroked, true, false)).not.toContain(
			'overflow-hidden',
		);
	});
});
