import { describe, expect, it } from 'vitest';

import { partNameToken } from './part-name';

describe('partNameToken', () => {
	it('keeps ordinary part names unchanged so derived ids stay stable', () => {
		for (const name of ['ppt/slides/slide1.xml', 'ppt/slideLayouts/slide_Layout-2.xml', 'a~b.c']) {
			expect(partNameToken(name)).toBe(name);
		}
	});

	it('leaves nothing that markup or CSS treats as syntax', () => {
		const token = partNameToken('ppt/slides/s</style><b x="1" y=\'2\'>&\\\n{}.xml');
		expect(token).toMatch(/^[A-Za-z0-9\-._~/%]+$/u);
	});

	it('encodes non-ASCII names as UTF-8', () => {
		expect(partNameToken('スlide 1.xml')).toBe('%E3%82%B9lide%201.xml');
	});

	it('never gives two different names the same token', () => {
		expect(partNameToken('a<b.xml')).not.toBe(partNameToken('a%3Cb.xml'));
	});

	it('survives a lone surrogate', () => {
		expect(partNameToken('a\ud800b')).toBe('a%EF%BF%BDb');
	});
});
