import { describe, expect, it } from 'vitest';

import { isSafeIdToken, partNameToken, safeIdToken } from './part-name';

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

describe('safeIdToken', () => {
	it('keeps ids the loader and the editor build', () => {
		for (const id of [
			'ppt/slides/slide1.xml-shape-0',
			'ppt/slides/slide%201.xml-shape-3',
			'layout-el-1763526000000-a1b2c3d',
			'tbl-merged-mhx1-abc123',
		]) {
			expect(isSafeIdToken(id)).toBe(true);
			expect(safeIdToken(id)).toBe(id);
		}
	});

	it('encodes any other text the way the loader encodes part names', () => {
		expect(isSafeIdToken('x"><img>')).toBe(false);
		expect(safeIdToken('ppt/slides/slide 1.xml-shape-0')).toBe('ppt/slides/slide%201.xml-shape-0');
		expect(safeIdToken('a</style>')).toMatch(/^[A-Za-z0-9\-._~/%]+$/u);
		expect(isSafeIdToken('')).toBe(false);
		expect(isSafeIdToken('50%')).toBe(false);
	});
});
