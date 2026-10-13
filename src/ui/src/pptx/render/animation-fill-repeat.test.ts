import { expect, it } from 'vitest';
import { resolveEffectTiming as coreResolveEffectTiming } from 'ooxml-core/pptx';
import { resolveEffectTiming } from './animation-fill-repeat';

it('uses the core effect timing implementation', () => {
	expect(resolveEffectTiming).toBe(coreResolveEffectTiming);
});
