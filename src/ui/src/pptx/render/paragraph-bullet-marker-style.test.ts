import { describe, expect, it } from 'vitest';

import { buildBulletMarkerStyle } from './paragraph-bullet-marker-style';

const bullet = (fontFamily?: string) => ({ marker: '+', isNumbered: false, fontFamily });

describe('buildBulletMarkerStyle typeface', () => {
	it('falls back to the base family of a weight or style variant named in a:buFont', () => {
		const style = buildBulletMarkerStyle(bullet('Lato Light Italic'), undefined, 1, undefined);
		expect(style.fontFamily).toMatch(/^"Lato Light Italic", "?Lato"?,/u);
		expect(style.fontWeight).toBe(300);
		expect(style.fontStyle).toBe('italic');
	});

	it('uses any other a:buFont as it is', () => {
		expect(buildBulletMarkerStyle(bullet('Wingdings'), undefined, 1, undefined).fontFamily).toBe(
			'Wingdings',
		);
		expect(buildBulletMarkerStyle(bullet('Lato'), undefined, 1, undefined).fontFamily).toBe('Lato');
	});
});
