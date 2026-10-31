// @vitest-environment jsdom
import { expect, it, vi } from 'vitest';
import { exportPageSvg } from './export-svg';
import { templateThumbnail } from './template-thumbnails';

vi.mock('./export-svg', async (importOriginal) => {
	const actual = await importOriginal<typeof import('./export-svg')>();
	return { ...actual, exportPageSvg: vi.fn(actual.exportPageSvg) };
});

it('retries a failed render while sharing pending and successful thumbnails', async () => {
	vi.mocked(exportPageSvg).mockImplementationOnce(() => {
		throw new Error('DOM is temporarily unavailable');
	});
	const first = templateThumbnail('blank');
	expect(templateThumbnail('blank')).toBe(first);
	expect(await first).toBeUndefined();
	const retry = templateThumbnail('blank');
	expect(retry).not.toBe(first);
	expect(await retry).toMatch(/^data:image\/svg\+xml;charset=utf-8,/);
	expect(templateThumbnail('blank')).toBe(retry);
});
