import { describe, expect, it } from 'vitest';

import { svgIdToken } from './svg-id';

describe('svgIdToken', () => {
	it('keeps letters, digits, dots and hyphens', () => {
		expect(svgIdToken('shape-1.a')).toBe('shape-1.a');
	});

	it.each(['ppt/slides/a%20b.xml-shape-0', 'a b', 'スライド', 'a"b<c>', 'x~1', '😀'])(
		'leaves only reference-safe characters for %s',
		(value) => {
			expect(svgIdToken(value)).toMatch(/^[A-Za-z0-9._-]+$/u);
		},
	);

	it('gives different values different tokens', () => {
		const values = ['a/b', 'a_b', 'a_2f_b', 'a%2fb', 'a b', 'a_20_b'];
		expect(new Set(values.map(svgIdToken)).size).toBe(values.length);
	});
});
