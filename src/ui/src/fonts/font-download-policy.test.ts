import { describe, expect, it, vi } from 'vitest';

import { effectiveFontDownloads, fontDownloadLock } from './font-download-policy';

describe('font download policy for Word, Excel and Visio', () => {
	it('lets the reader decide when the host set nothing', () => {
		expect(effectiveFontDownloads(undefined, false)).toBe(false);
		expect(effectiveFontDownloads(undefined, true)).toEqual({
			baseUrl: 'https://fonts.googleapis.com/css2',
		});
		expect(fontDownloadLock(undefined)).toBeUndefined();
	});

	it('lets a host value win and lock the toggle', () => {
		const mirror = { baseUrl: 'https://fonts.example.com/css2' };
		expect(effectiveFontDownloads(mirror, false)).toBe(mirror);
		expect(effectiveFontDownloads(false, true)).toBe(false);
		expect(fontDownloadLock(mirror)).toBe(true);
		expect(fontDownloadLock(false)).toBe(false);
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
		try {
			expect(fontDownloadLock({ baseUrl: 'nope' })).toBe(false);
		} finally {
			warn.mockRestore();
		}
	});
});
