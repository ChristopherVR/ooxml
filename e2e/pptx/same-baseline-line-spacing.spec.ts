/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { readFile, writeFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import { elementWithText, fixture, loadDeckAt, ribbonTab, selectElement } from './support/deck';
import { downloadBytes } from './support/exports';
import { acrossFrameworks } from './support/parity';

test(
	'a whole-body spacing choice equal to the first paragraph edits inherited paragraphs and supports undo',
	{ tag: '@cross-binding' },
	async ({ browser }, info) => {
		test.slow();
		const zip = await JSZip.loadAsync(await readFile(fixture('linked-textbox.pptx')));
		const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		const body =
			'<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="200000"/></a:lnSpc></a:lvl1pPr></a:lstStyle><a:p><a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="2400"/><a:t>First spacing paragraph</a:t></a:r></a:p><a:p><a:r><a:rPr sz="2400"/><a:t>Second spacing paragraph</a:t></a:r></a:p></p:txBody>';
		zip.file('ppt/slides/slide1.xml', xml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/u, body));
		const source = info.outputPath('same-baseline-source.pptx');
		await writeFile(source, await zip.generateAsync({ type: 'nodebuffer' }));
		await acrossFrameworks(
			browser,
			info,
			async (page, url) => {
				await loadDeckAt(page, url, source);
				await ribbonTab(page, 'Home').click();
				await selectElement(page, elementWithText(page, 'First spacing paragraph'));
				const control = page.locator('[data-ribbon-control="home.paragraph.lineSpacing"]').first();
				await control.getByRole('combobox').click();
				await page.getByRole('option', { name: '1.5', exact: true }).click();
				async function secondProperties() {
					const bytes = await downloadBytes(await savePptxViaBackstage(page));
					const saved = await JSZip.loadAsync(bytes);
					const slide = await saved.file('ppt/slides/slide1.xml')!.async('string');
					const paragraphs = slide.match(/<a:p>[\s\S]*?<\/a:p>/gu)!;
					const second = paragraphs.find((p) => p.includes('Second spacing paragraph'))!;
					expect(second).toBeDefined();
					return second.match(/<a:pPr>[\s\S]*?<\/a:pPr>/u)?.[0] ?? '';
				}
				expect(await secondProperties()).toContain('<a:spcPct val="150000"');
				// Save can return focus to the ribbon select, where keyboard undo is
				// ignored by document shortcuts. Use the document history control.
				await page.getByRole('button', { name: 'Undo', exact: true }).first().click();
				expect(await secondProperties()).not.toContain('<a:lnSpc>');
				await page.getByRole('button', { name: 'Redo', exact: true }).first().click();
				expect(await secondProperties()).toContain('<a:spcPct val="150000"');
				return true;
			},
			{ viewport: { width: 2600, height: 1000 } },
		);
	},
);
