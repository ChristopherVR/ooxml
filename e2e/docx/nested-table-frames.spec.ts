import { expect, test } from '@playwright/test';
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileInput } from './helpers';
for (const name of [
	'short',
	'two-rows',
	'centered',
	'split',
	'align-top',
	'align-center',
	'align-bottom',
	'border-half',
	'border-two',
	'border-three-margin',
	'margin',
	'unequal-split',
	'mixed',
	'mixed-margin',
	'mixed-centered',
	'mixed-two-rows',
	'mixed-split',
	'mixed-outer-border',
	'mixed-deep',
])
	test(`nested table frames ${name} survive save and reopen`, async ({ page }) => {
		await page.goto('/?framework=vanilla');
		const editor = page.locator('docx-editor');
		const input = editor.locator('.ProseMirror').first();
		await (
			await fileInput(page)
		).setInputFiles({
			name: name + '.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: readFileSync(
				new URL(
					'../../src/core/docx/__fixtures__/nested-table-frames/nested-frame-' + name + '.docx',
					import.meta.url,
				),
			),
		});
		await expect(input).toContainText('Outer cell');
		const print = async () => {
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			await editor
				.getByRole('combobox', { name: 'Layout view', exact: true })
				.selectOption('print');
		};
		await print();
		const frames = editor.locator('.dve-print-nested-table');
		await expect(frames).toHaveCount(name.endsWith('split') || name === 'mixed-deep' ? 2 : 1);
		const check = async () => {
			const first = frames.first().locator('.dve-print-cell').first();
			await expect(first).toHaveCSS('background-color', 'rgb(255, 255, 0)');
			await expect(first).toHaveCSS('border-left-color', 'rgb(255, 0, 0)');
			expect(
				await first.evaluate((node) => parseFloat((node as HTMLElement).style.borderLeftWidth)),
			).toBeCloseTo(
				((name === 'border-half'
					? 0.5
					: name === 'border-two'
						? 2
						: name === 'border-three-margin'
							? 3
							: 1) *
					4) /
					3,
				1,
			);
			await expect(frames.first()).toHaveText('');
			const layer = frames
				.first()
				.locator('xpath=ancestor::*[contains(@class,"dve-print-cell-frame-layer")][1]');
			await expect(layer).toHaveCSS('overflow-x', 'visible');
			await expect(layer).toHaveCSS('overflow-y', 'clip');
			await expect(layer).toHaveCSS('pointer-events', 'none');
			expect(await layer.evaluate((node) => node.closest('.dve-print-cell'))).toBeNull();
			// Read the frame and enclosing cell from one paint, including asynchronous reopen.
			type Bounds = {
				frameBounds: { x: number; y: number; width: number; height: number };
				bodyCell: { x: number; y: number; width: number; height: number };
			};
			const snapshot = { bounds: null as Bounds | null };
			await expect
				.poll(async () => {
					snapshot.bounds = await editor.evaluate((root) => {
						const rendered = root.shadowRoot ?? root;
						const frame = rendered.querySelector('.dve-print-nested-table');
						const body = rendered.querySelector('.dve-print-page .dve-print-row > .dve-print-cell');
						if (!frame || !body) return null;
						const a = frame.getBoundingClientRect(),
							b = body.getBoundingClientRect();
						if (!a.width || !b.width) return null;
						return {
							frameBounds: { x: a.x, y: a.y, width: a.width, height: a.height },
							bodyCell: { x: b.x, y: b.y, width: b.width, height: b.height },
						};
					});
					return snapshot.bounds !== null;
				})
				.toBe(true);
			const { frameBounds, bodyCell } = snapshot.bounds!;
			const evidenceDirectory = process.env.WORD_FRAME_PAINT_EVIDENCE_DIRECTORY;
			if (name === 'mixed' && evidenceDirectory) {
				mkdirSync(evidenceDirectory, { recursive: true });
				const ink = await editor
					.locator('.dve-print-block[data-nested="true"]')
					.filter({ hasText: 'T1R0C0L00' })
					.boundingBox();
				await page.screenshot({ path: join(evidenceDirectory, 'word-mixed-frame.png') });
				const originalHeight = await layer.evaluate((node) => (node as HTMLElement).style.height);
				try {
					await layer.evaluate((node) => {
						(node as HTMLElement).style.height = '8px';
					});
					await page.screenshot({ path: join(evidenceDirectory, 'word-mixed-frame-clipped.png') });
				} finally {
					await layer.evaluate((node, height) => {
						(node as HTMLElement).style.height = height;
					}, originalHeight);
				}

				writeFileSync(
					join(evidenceDirectory, 'word-mixed-frame-geometry.json'),
					JSON.stringify(
						{ frame: frameBounds, body: bodyCell, ink, viewport: page.viewportSize() },
						null,
						2,
					),
					'utf8',
				);
			}

			if (name !== 'centered' && name !== 'mixed-centered')
				expect(frameBounds!.x + frameBounds!.width).toBeGreaterThan(
					bodyCell!.x + bodyCell!.width + 0.4,
				);

			if (name === 'mixed-outer-border' || name === 'mixed-deep') {
				const paragraph = editor
					.locator('.dve-print-block[data-nested="true"]')
					.filter({ hasText: 'T1R0C0L00' });
				const ink = await paragraph.boundingBox();
				expect(ink!.x - frameBounds!.x).toBeCloseTo(4 / 3, 1);
				expect(ink!.y - frameBounds!.y).toBeCloseTo(
					((name === 'mixed-deep' ? 16.5 : 2) * 4) / 3,
					1,
				);
				if (name === 'mixed-deep') {
					const deepFrame = await frames.nth(1).boundingBox();
					const deepInk = await editor
						.locator('.dve-print-block[data-nested="true"]')
						.filter({ hasText: 'D0' })
						.boundingBox();
					expect(deepFrame!.x - frameBounds!.x).toBeCloseTo(4 / 3, 1);
					expect(deepFrame!.y - frameBounds!.y).toBeCloseTo(8 / 3, 1);
					expect(deepInk!.x - deepFrame!.x).toBeCloseTo(4 / 3, 1);
					expect(deepInk!.y - deepFrame!.y).toBeCloseTo(8 / 3, 1);
				}
			}
			if (name.startsWith('mixed')) {
				const firstGeometry = await frames
					.first()
					.locator('.dve-print-cell')
					.first()
					.evaluate((node) => ({
						left: parseFloat((node as HTMLElement).style.left),
						width: parseFloat((node as HTMLElement).style.width),
					}));
				const secondGeometry = await frames
					.first()
					.locator('.dve-print-cell')
					.nth(1)
					.evaluate((node) => ({
						left: parseFloat((node as HTMLElement).style.left),
						width: parseFloat((node as HTMLElement).style.width),
					}));
				expect(secondGeometry.left).toBeCloseTo(firstGeometry.left + firstGeometry.width, 1);
				expect(
					parseFloat(await frames.first().evaluate((node) => (node as HTMLElement).style.width)),
				).toBeCloseTo(((name === 'mixed-centered' ? 122 : 162) * 4) / 3, 1);
			}

			if (name === 'unequal-split') {
				const continued = editor.locator('.dve-print-page').nth(1);
				await expect(continued).toContainText('T1R0C0L19');
				await expect(continued).toContainText('T1R0C1L17');
				await expect(continued).not.toContainText('T1R0C1L18');
				const long = continued
					.locator('.dve-print-block[data-nested="true"]')
					.filter({ hasText: 'T1R0C0L19' });
				const box = await long.boundingBox();
				const cell = await continued.locator('.dve-print-cell').first().boundingBox();
				expect(box!.y + box!.height).toBeLessThanOrEqual(cell!.y + cell!.height);
				await expect(continued.locator('.dve-print-footnote')).toContainText('N0L02');
			}

			if (name.startsWith('align-')) {
				const paragraph = editor
					.locator('.dve-print-block[data-nested="true"]')
					.filter({ hasText: 'T1R0C0L00' });
				const paragraphY = await paragraph.evaluate((node) =>
					parseFloat((node as HTMLElement).style.top),
				);
				const frameY = await frames
					.first()
					.evaluate((node) => parseFloat((node as HTMLElement).style.top));
				expect(paragraphY - frameY).toBeCloseTo(
					name === 'align-center' ? 38 / 3 : name === 'align-bottom' ? 24 : 4 / 3,
					1,
				);
			}

			await expect(
				editor.locator('.dve-print-block[data-nested="true"]').filter({ hasText: 'T1R0C0L00' }),
			).toHaveCount(1);
		};
		await check();
		const bytes = await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'reopened.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(bytes),
		});
		await expect(input).toContainText('Outer cell');
		await print();
		await check();
	});
