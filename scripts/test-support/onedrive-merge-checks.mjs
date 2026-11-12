import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { fileCommand } from '../suite-check-init.mjs';
const require = createRequire(new URL('../../e2e/docx/package.json', import.meta.url));
const { expect } = require('@playwright/test');
const { inspectDocx, setDocxRunText, readXlsxRange, setXlsxCells } =
	await import('../../src/core/dist/automation/index.mjs');
const { inspectPresentationText, setPresentationRunText } =
	await import('../../src/core/dist/pptx/index.mjs');
const { editVsdx, parseVsdx } = await import('../../src/core/dist/visio/index.mjs');

/** Two independently changed saved files must become one cloud item, without recovery UI. */
export async function checkConcurrentMerges({ page, cloud, openCloud, edit }) {
	for (const kind of ['docx', 'xlsx', 'pptx', 'vsdx']) {
		await openCloud(kind);
		await edit(kind, 'Local collaboration');
		const remote = cloud.files.get(kind);
		if (kind === 'docx') {
			const paragraph = (await inspectDocx(remote.bytes)).blocks.find(
				(block) => block.type === 'paragraph',
			);
			remote.bytes = Buffer.from(
				(
					await setDocxRunText(
						remote.bytes,
						paragraph.id,
						0,
						`Remote collaboration: ${paragraph.runs[0].text}`,
					)
				).bytes,
			);
		} else if (kind === 'xlsx') {
			remote.bytes = Buffer.from(
				(await setXlsxCells(remote.bytes, 0, [{ address: 'B1', input: 'Remote collaboration' }]))
					.bytes,
			);
		} else if (kind === 'pptx') {
			const element = (await inspectPresentationText(remote.bytes)).slides[0].elements.find(
				(shape) => shape.text,
			);
			assert.ok(element);
			remote.bytes = Buffer.from(
				await setPresentationRunText(remote.bytes, 0, element.id, 0, 'Remote collaboration'),
			);
		} else {
			remote.bytes = Buffer.from(
				(
					await editVsdx(remote.bytes, [
						{ type: 'format-shape', pageId: '0', shapeId: '1', fillColor: '#00FF00' },
					])
				).bytes,
			);
		}
		remote.revision++;
		const writes = cloud.writes.length;
		await fileCommand(page, 'save');
		await expect.poll(() => cloud.writes.length).toBe(writes + 1);
		await expect(page.locator('#save-state')).toHaveText('Saved to OneDrive');
		await expect(page.locator('#onedrive-conflict')).toHaveCount(0);
		if (kind === 'docx') {
			const text = JSON.stringify(await inspectDocx(remote.bytes));
			assert.match(text, /Local collaboration/u);
			assert.match(text, /Remote collaboration/u);
			await expect(page.locator('.document-host:not([hidden]) docx-editor')).toContainText(
				'Remote collaboration',
			);
		} else if (kind === 'xlsx') {
			assert.equal(
				(await readXlsxRange(remote.bytes, 0, 'A1')).cells[0].value,
				'Local collaboration',
			);
			assert.equal(
				(await readXlsxRange(remote.bytes, 0, 'B1')).cells[0].value,
				'Remote collaboration',
			);
			await expect
				.poll(() =>
					page
						.locator('.document-host:not([hidden]) xlsx-editor')
						.evaluate((editor) => editor.workbook.sheets[0].rows.get(0)?.get(1)?.value),
				)
				.toBe('Remote collaboration');
		} else if (kind === 'pptx') {
			const presentation = await inspectPresentationText(remote.bytes);
			assert.equal(presentation.slides.length, 3);
			assert.match(JSON.stringify(presentation), /Remote collaboration/u);
			await expect(
				page.locator('.document-host:not([hidden]) [aria-label="Go to slide 3"]'),
			).toBeVisible();
		} else {
			const text = JSON.stringify(await parseVsdx(remote.bytes));
			assert.match(text, /Local collaboration/u);
			assert.match(text, /00ff00/iu);
			await expect(page.locator('.document-host:not([hidden]) visio-viewer')).toContainText(
				'Local collaboration',
			);
		}
		console.log(
			`${kind}: concurrent local and remote edits combined into the same saved file and editor`,
		);
	}
}
