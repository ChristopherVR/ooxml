import { describe, expect, it } from 'vitest';

import type { TextStyle, XmlObject } from '../../types';
import { PptxHandlerRuntime } from './PptxHandlerRuntimeImplementation';

/** The real `applyTextRunEffects`, with the two colour helpers it calls stubbed. */
function apply(effectList: XmlObject, alpha: number | undefined): TextStyle {
	const style: TextStyle = {};
	const host = {
		parseColor: () => '#000000',
		extractColorOpacity: () => alpha,
	};
	(
		PptxHandlerRuntime.prototype as unknown as {
			applyTextRunEffects(this: unknown, style: TextStyle, list: XmlObject): void;
		}
	).applyTextRunEffects.call(host, style, effectList);
	return style;
}

describe('a text run outer shadow without a:alpha', () => {
	it('is opaque (ECMA-376), not the renderer default of 0.5', () => {
		expect(apply({ 'a:outerShdw': { '@_blurRad': '152400' } }, undefined).textShadowOpacity).toBe(
			1,
		);
	});

	it('keeps an authored alpha', () => {
		expect(apply({ 'a:outerShdw': { '@_blurRad': '152400' } }, 0.4).textShadowOpacity).toBe(0.4);
	});
});
