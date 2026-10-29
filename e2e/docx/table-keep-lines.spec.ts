import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { buildXml, parseXml } from '../../src/core/xml/index';
import { fileInput } from './helpers';
const root = new URL('../../src/core/docx/__fixtures__/table-keep-lines/', import.meta.url);
const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const evidence = JSON.parse(readFileSync(new URL('native-measurements.json', root), 'utf8')) as {
	cases: {
		name: string;
		measurements: {
			borders: { rectPt: number[] }[];
			lines: { text: string; baselinePt: number }[];
		}[];
	}[];
};
const tableProps = (xml: string) => buildXml(parseXml(xml).getElementsByTagNameNS(w, 'tblPr')[0]!);
for (const profile of ['partial', 'fresh', 'oversize', 'consumed', 'mirror-fresh'])
	test(
		'table-cell keep-lines controls, native bands, history and fresh reopen: ' + profile,
		async ({ page }) => {
			const on = evidence.cases.find((c) => c.name === 'table-keep-' + profile + '-on')!;
			const off = evidence.cases.find((c) => c.name === 'table-keep-' + profile + '-off')!;
			const original = readFileSync(new URL(on.name + '.docx', root)),
				zip = await JSZip.loadAsync(original);
			await page.goto('/?framework=vanilla');
			const editor = page.locator('docx-editor'),
				source = editor.locator('.ProseMirror').first();
			await (
				await fileInput(page)
			).setInputFiles({
				name: 'keeping.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer: original,
			});
			await expect(source).toContainText('T1R0C0LK00');
			const print = async () => {
				await editor.getByRole('tab', { name: 'View', exact: true }).click();
				await editor
					.getByRole('combobox', { name: 'Layout view', exact: true })
					.selectOption('print');
				await expect(editor.locator('.dve-print-pages').first()).toBeVisible();
			};
			await print();
			const check = async (native: typeof on) => {
				const pages = editor.locator('.dve-print-page');
				await expect(pages).toHaveCount(native.measurements.length);
				await expect
					.poll(async () =>
						(await pages.allTextContents()).map((text) =>
							(text.match(/T[01]R\d+C\d+L[A-Z]\d{2}/g) ?? []).sort(),
						),
					)
					.toEqual(
						native.measurements.map((sample) =>
							sample.lines
								.map((line) => line.text.trim())
								.filter((text) => /^T[01]/.test(text))
								.sort(),
						),
					);
				for (const [index, sample] of native.measurements.entries()) {
					const rendered = pages.nth(index),
						frame = rendered.locator('.dve-print-table');
					await expect(frame).toHaveCount(sample.borders.length ? 1 : 0);
					expect(
						((await rendered.textContent())?.match(/T[01]R\d+C\d+L[A-Z]\d{2}/g) ?? []).sort(),
					).toEqual(
						sample.lines
							.map((line) => line.text.trim())
							.filter((text) => /^T[01]/.test(text))
							.sort(),
					);
					if (sample.borders.length) {
						const top = Math.min(...sample.borders.map((b) => b.rectPt[1]!));
						const bottom = Math.max(...sample.borders.map((b) => b.rectPt[3]!));
						const nativeFirst = sample.lines.find((line) => /^T1/.test(line.text))!;
						// A pending relayout can replace the page after reopening: read both boxes together.
						await expect(async () => {
							const geometry = await frame.evaluate((node) => {
								const first = node.querySelector('.dve-print-cell-content > .dve-print-block')!;
								const bounds = node.getBoundingClientRect();
								return {
									height: bounds.height,
									offset: first.getBoundingClientRect().y - bounds.y,
								};
							});
							expect(Math.abs(geometry.height * 0.75 - (bottom - top))).toBeLessThan(0.2);
							expect(
								Math.abs(geometry.offset * 0.75 - (nativeFirst.baselinePt - 9.6 - top)),
							).toBeLessThan(0.15);
						}).toPass({ timeout: 5000 });
					}
				}
			};
			await check(on);
			const save = async () =>
				Buffer.from(
					await editor.evaluate(async (node) =>
						Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
					),
				);
			const noOp = await JSZip.loadAsync(await save());
			for (const [path, file] of Object.entries(zip.files))
				if (!file.dir)
					expect(await noOp.file(path)!.async('uint8array')).toEqual(
						await file.async('uint8array'),
					);
			const target = editor
				.locator('.dve-print-cell-content > .dve-print-block')
				.filter({ hasText: 'T1R0C0LK00' })
				.first();
			const toggle = async (keep: boolean) => {
				await target.click();
				await editor.getByRole('tab', { name: 'Home', exact: true }).click();
				await editor.getByRole('button', { name: 'Paragraph settings', exact: true }).click();
				const dialog = editor.locator('.dve-paragraph-dialog');
				await dialog
					.getByRole('checkbox', { name: 'Keep lines together', exact: true })
					.setChecked(keep);
				await dialog.getByRole('button', { name: 'OK', exact: true }).click();
			};
			await toggle(false);
			await check(off);
			const offXml = await (
				await JSZip.loadAsync(await save())
			)
				.file('word/document.xml')!
				.async('string');
			const changed = Array.from(parseXml(offXml).getElementsByTagNameNS(w, 'p')).find((p) =>
				p.textContent?.includes('T1R0C0LK00'),
			)!;
			expect(changed.getElementsByTagNameNS(w, 'keepLines')[0]!.getAttributeNS(w, 'val')).toBe('0');
			await target.click();
			await page.keyboard.press('Control+z');
			await expect
				.poll(async () =>
					(await editor.locator('.dve-print-page').allTextContents()).map(
						(text) => text.match(/T1R0C0LK\d{2}/g) ?? [],
					),
				)
				.toEqual(
					on.measurements.map((sample) =>
						sample.lines.map((line) => line.text.trim()).filter((text) => /^T1R0C0LK/.test(text)),
					),
				);
			await check(on);
			await page.keyboard.press('Control+Shift+z');
			await check(off);
			await toggle(true);
			await check(on);
			await target.click({ position: { x: 10, y: 2 } });
			await page.keyboard.press('Home');
			await page.keyboard.type('Ed');
			await expect(source.locator('p').filter({ hasText: 'T1R0C0LK00' })).toContainText(
				'EdT1R0C0LK00',
			);
			await expect(target).toContainText('EdT1R0C0LK00');
			const bytes = await save(),
				edited = await JSZip.loadAsync(bytes);
			expect(tableProps(await edited.file('word/document.xml')!.async('string'))).toBe(
				tableProps(await zip.file('word/document.xml')!.async('string')),
			);
			for (const [path, file] of Object.entries(zip.files))
				// docProps/core.xml carries the save's last-modified stamp.
				if (!file.dir && path !== 'word/document.xml' && path !== 'docProps/core.xml')
					expect(await edited.file(path)!.async('uint8array')).toEqual(
						await file.async('uint8array'),
					);
			const oldSource = await source.elementHandle(),
				oldPage = await editor.locator('.dve-print-page').last().elementHandle();
			await (
				await fileInput(page)
			).setInputFiles({
				name: 'reopen.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer: bytes,
			});
			await expect.poll(() => oldSource!.evaluate((node) => node.isConnected)).toBe(false);
			await expect.poll(() => oldPage!.evaluate((node) => node.isConnected)).toBe(false);
			await oldSource!.dispose();
			await oldPage!.dispose();
			await expect(source.locator('p').filter({ hasText: 'T1R0C0LK00' })).toContainText(
				'EdT1R0C0LK00',
			);
			await print();
			await expect(target).toContainText('EdT1R0C0LK00');
			await check(on);
		},
	);
