import { readFileSync } from 'node:fs';
import path from 'node:path';
import { expect, it } from 'vitest';
import { createEditSession, loadXlsx } from 'ooxml-core/xlsx';
import { buildPrintHtml } from './print-pages';

const fixture = () =>
	new Uint8Array(
		readFileSync(
			path.join(
				import.meta.dirname,
				'..',
				'..',
				'..',
				'core',
				'xlsx',
				'__fixtures__',
				'excel-header-images.xlsx',
			),
		),
	);

it('prints header pictures, page-number offsets, theme colours and double underline', async () => {
	const wb = await loadXlsx(fixture());
	const html = buildPrintHtml(wb, 0, 'Book.xlsx');
	// Page 1 shows the first-page header, whose left section holds a picture.
	expect(html.match(/<img src="data:image\/png;base64,/g)).toHaveLength(1);
	expect(html).toContain('width:40px;height:20px');
	const edit = createEditSession(wb);
	edit.setPageSetup(0, {
		differentFirst: false,
		header: '&C&K05+000Page &P+2&E&Xx',
	});
	const next = buildPrintHtml(wb, 0, 'Book.xlsx');
	expect(next).toContain('Page 3');
	expect(next).toMatch(/color:#[0-9A-F]{6}/);
	expect(next).toContain('text-decoration-style:double');
	expect(next).toContain('vertical-align:super');
});
