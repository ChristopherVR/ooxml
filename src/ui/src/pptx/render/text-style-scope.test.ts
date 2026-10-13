import { expect, it } from 'vitest';
import * as coreScope from 'ooxml-core/pptx/editor/render/text-style-scope';
import * as uiScope from './text-style-scope';
import * as coreContent from 'ooxml-core/pptx/editor/render/text-content-style';
import { isContentSegment, TEXT_DECORATION_FLAGS } from './selection-format-state';

it('retains core implementations through the shared UI compatibility paths', () => {
	for (const name of Object.keys(coreScope) as (keyof typeof coreScope)[]) {
		expect(uiScope[name]).toBe(coreScope[name]);
	}
	expect(isContentSegment).toBe(coreContent.isContentSegment);
	expect(TEXT_DECORATION_FLAGS).toBe(coreContent.TEXT_DECORATION_FLAGS);
});
