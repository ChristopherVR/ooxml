import { describe, expect, it } from 'vitest';

import { ensureType } from './package';
import { mediaEntry } from './schema';

const TYPES =
	'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';

describe('media a peer shares', () => {
	it('keeps a media part of a well-formed type', () => {
		expect(
			mediaEntry({ path: 'visio/media/image1.png', hash: 'h', ct: 'image/png', d: true }),
		).toEqual({
			path: 'visio/media/image1.png',
			hash: 'h',
			ct: 'image/png',
			d: true,
		});
	});

	it('refuses a path outside a media folder or a malformed type', () => {
		for (const entry of [
			{ path: 'visio/document.xml', ct: 'image/png' },
			{ path: 'visio/../media/x.png', ct: 'image/png' },
			{ path: '/visio/media/x.png', ct: 'image/png' },
			{ path: 'visio\\media\\x.png', ct: 'image/png' },
			{ path: 'visio/media/x.png', ct: 'image/png"/><Override PartName="/x"' },
		]) {
			expect(mediaEntry({ ...entry, hash: 'h' })).toBeUndefined();
		}
	});

	it('escapes what it writes into the content types part', () => {
		const xml = ensureType(TYPES, 'visio/media/a"b.png', { type: 'image/png', byDefault: false });
		expect(xml).toContain('PartName="/visio/media/a&#34;b.png"');
	});
});
