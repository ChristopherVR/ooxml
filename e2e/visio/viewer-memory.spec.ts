import JSZip from 'jszip';
import { expect, test } from '@playwright/test';
import { heapReport, openAndCloseEditor } from '../../scripts/heap-snapshot.mjs';
import { openDemo } from './demo-page';
import { createVsdxFixture } from './fixture.mjs';

/** The fixture's shape with a glow: a group drawn through `filter: url(#id)`. */
async function glowingShape(): Promise<Buffer> {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Glowing'));
	const page = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		page.replace(
			'<Cell N="FillForegnd" V="#DAEFE6"/>',
			'<Cell N="FillForegnd" V="#DAEFE6"/><Cell N="GlowColor" V="#C00000"/><Cell N="GlowSize" V="0.1"/>',
		),
	);
	return Buffer.from(await zip.generateAsync({ type: 'uint8array' }));
}

// Glow, reflection and soft shadows paint groups through `url(#id)` filters and masks. Chromium's
// style cache kept a closed viewer alive through them until some later stylesheet change.
const CYCLES = 3;
const PROBE = /^<visio-viewer[^>]*data-leak-probe/;

test('viewers opened and closed repeatedly are collected', async ({ page }) => {
	test.setTimeout(240_000);
	await openDemo(page);
	const session = await page.context().newCDPSession(page);
	const start = await heapReport(session, PROBE);
	const bytes = [...(await glowingShape())];
	let references = 0;
	for (let cycle = 1; cycle <= CYCLES; cycle++) {
		references += await page.evaluate(openAndCloseEditor, {
			tag: 'visio-viewer',
			bytes,
			fileName: 'glow.vsdx',
		});
		const report = await heapReport(session, PROBE);
		test.info().annotations.push({
			type: 'heap',
			description: `glow.vsdx close ${cycle}: ${report.count} closed viewers alive, ${report.detachedBytes} detached bytes, heap ${report.heapBytes - start.heapBytes} bytes over the start`,
		});
		expect.soft(report.count, `close ${cycle}`).toBe(0);
	}
	expect(references).toBeGreaterThan(0);
});
