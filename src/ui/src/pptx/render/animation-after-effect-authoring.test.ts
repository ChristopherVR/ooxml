import { expect, it } from 'vitest';
import { setAfterAnimation as coreSetAfterAnimation } from 'ooxml-core/pptx';
import { setAfterAnimation } from './animation-after-effect-authoring';

it('keeps the shared UI authoring entry as the core implementation', () => {
	expect(setAfterAnimation).toBe(coreSetAfterAnimation);
});
