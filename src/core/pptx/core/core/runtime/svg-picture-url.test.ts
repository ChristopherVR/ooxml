import { readFile } from 'node:fs/promises';
import { join } from 'node:path';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../../index';

const fixture = join(import.meta.dirname, '../../../__tests__/fixtures/svg-picture-crop.pptx');

describe('picture URLs', () => {
	it('never give an SVG picture a blob URL of the page origin', async () => {
		// Node has URL.createObjectURL, so raster pictures get blob URLs here as in a browser.
		expect(typeof URL.createObjectURL).toBe('function');
		const zip = await JSZip.loadAsync(await readFile(fixture));
		const svg =
			'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><rect width="10" height="10"/></svg>';
		zip.file('ppt/media/picture.svg', svg);
		const handler = new PptxHandler();
		await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));

		const svgUrl = await handler.getImageData('ppt/media/picture.svg');
		expect(svgUrl?.startsWith('data:image/svg+xml;base64,')).toBe(true);
		const rendered = atob(svgUrl!.slice(svgUrl!.indexOf(',') + 1));
		expect(rendered).toContain('<rect width="10" height="10"/>');

		const pngUrl = await handler.getImageData('ppt/media/image1.png');
		expect(pngUrl?.startsWith('blob:')).toBe(true);
	});
});
