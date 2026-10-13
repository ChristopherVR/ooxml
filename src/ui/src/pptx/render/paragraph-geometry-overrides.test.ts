import { expect, it } from 'vitest';
import { resolveParagraphGeometryOverrides as core } from 'ooxml-core/pptx/editor/render/paragraph-geometry-overrides';
import { resolveParagraphGeometryOverrides as shared } from './paragraph-geometry-overrides';

it('retains the shared UI compatibility export as the core implementation', () => {
	expect(shared).toBe(core);
});
