import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { parseXml, buildXml } from '../../src/core/xml/index';
import { fileInput } from './helpers';
const twoRoot = new URL('../../src/core/docx/__fixtures__/nested-multi-row/', import.meta.url);
const evidence = JSON.parse(readFileSync(new URL('native-measurements.json', twoRoot), 'utf8')) as {
	cases: {
		name: string;
		measurements: {
			borders: { rectPt: number[] }[];
			lines: { text: string; baselinePt: number }[];
		}[];
	}[];
};
const nested = (xml: string) =>
	buildXml(
		Array.from(
			parseXml(xml).getElementsByTagNameNS(
				'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
				'tbl',
			),
		)[1]!,
	);
const threeRoot = new URL('../../src/core/docx/__fixtures__/nested-three-row/', import.meta.url);
const threeEvidence = JSON.parse(
	readFileSync(new URL('native-measurements.json', threeRoot), 'utf8'),
) as typeof evidence;
const structuralRoot = new URL(
	'../../src/core/docx/__fixtures__/nested-structural-rows/',
	import.meta.url,
);
const structuralEvidence = JSON.parse(
	readFileSync(new URL('native-measurements.json', structuralRoot), 'utf8'),
) as typeof evidence;
for (const { native, root } of [
	...evidence.cases.map((native) => ({ native, root: twoRoot })),
	...threeEvidence.cases.map((native) => ({ native, root: threeRoot })),
	...structuralEvidence.cases
		.filter((native) => /-(4-first|8-short|8-all-long|8-minimal)$/.test(native.name))
		.map((native) => ({ native, root: structuralRoot })),
])
	test(
		(native.name.startsWith('nested-structural')
			? 'structural'
			: native.name.startsWith('nested-three')
				? 'three'
				: 'two') +
			'-row nested fragments preserve native bands and guarded editing: ' +
			native.name,
		async ({ page }) => {
			await page.goto('/?framework=vanilla');
			const editor = page.locator('docx-editor'),
				input = editor.locator('.ProseMirror').first();
			const original = readFileSync(new URL(native.name + '.docx', root)),
				zip = await JSZip.loadAsync(original);
			const upload = async (bytes: Buffer) => {
				await (
					await fileInput(page)
				).setInputFiles({
					name: 'multi.docx',
					mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
					buffer: bytes,
				});
				await expect(input).toContainText('T2R0C0L00');
			};
			await upload(original);
			const print = async () => {
				await editor.getByRole('tab', { name: 'View', exact: true }).click();
				await editor
					.getByRole('combobox', { name: 'Layout view', exact: true })
					.selectOption('print');
			};
			await print();
			const check = async () => {
				const pages = editor.locator('.dve-print-page');
				await expect(pages).toHaveCount(native.measurements.length);
				for (const [index, sample] of native.measurements.entries()) {
					// Scheduled relayout can replace pages between awaited geometry reads.
					// Capture markers, frame and ink together from one connected visible page.
					const readSnapshot = () =>
						pages.nth(index).evaluate((node) => {
							const frames = node.querySelectorAll<HTMLElement>('.dve-print-nested-table');
							const frame = frames[0],
								ink = node.querySelector<HTMLElement>('.dve-print-block[data-read-only="true"]');
							if (!node.isConnected || frames.length !== 1 || !frame || !ink) return null;
							const bounds = frame.getBoundingClientRect(),
								inkBounds = ink.getBoundingClientRect();
							if (!bounds.width || !bounds.height || !inkBounds.width || !inkBounds.height)
								return null;
							const rows = frame.querySelectorAll<HTMLElement>(':scope > .dve-print-row');
							const first = rows[0]?.querySelector<HTMLElement>('.dve-print-cell'),
								last = rows[rows.length - 1]?.querySelector<HTMLElement>('.dve-print-cell');
							if (!first || !last) return null;
							return {
								tokens: (node.textContent?.match(/T[01]R\dC\dL\d{2}/g) ?? []).sort(),
								frameText: frame.textContent ?? '',
								pointerEvents: getComputedStyle(frame).pointerEvents,
								height: bounds.height,
								ink: inkBounds.y - bounds.y,
								topBorder: parseFloat(first.style.borderTopWidth),
								bottomBorder: parseFloat(last.style.borderBottomWidth),
							};
						});
					let snapshot = await readSnapshot();
					await expect
						.poll(async () => {
							snapshot = await readSnapshot();
							return snapshot?.tokens;
						})
						.toEqual(
							sample.lines
								.filter((line) => /^T[01]/.test(line.text))
								.map((line) => line.text.trim())
								.sort(),
						);
					if (!snapshot) throw new Error('Missing connected visible print snapshot');
					expect(snapshot.frameText).toBe('');
					expect(snapshot.pointerEvents).toBe('none');
					const top = Math.min(...sample.borders.map((b) => b.rectPt[1]!)),
						bottom = Math.max(...sample.borders.map((b) => b.rectPt[3]!)),
						firstNative = sample.lines.find((line) => /^T1/.test(line.text))!;
					expect(Math.abs(snapshot.height * 0.75 - (bottom - top))).toBeLessThan(0.2);
					expect(Math.abs(snapshot.ink * 0.75 - (firstNative.baselinePt - 9.6 - top))).toBeLessThan(
						0.15,
					);
					expect(snapshot.topBorder).toBeCloseTo(8 / 3);
					expect(snapshot.bottomBorder).toBeCloseTo(2 / 3);
				}
			};
			await check();
			const save = async () =>
				Buffer.from(
					await editor.evaluate(async (node) =>
						Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
					),
				);
			const unchanged = await JSZip.loadAsync(await save());
			expect(await unchanged.file('word/document.xml')!.async('string')).toBe(
				await zip.file('word/document.xml')!.async('string'),
			);
			const readonly = editor.locator('.dve-print-block[data-read-only="true"]').first();
			await readonly.click();
			await page.keyboard.type('Oops');
			await expect(input).not.toContainText('Oops');
			const outer = editor
				.locator('.dve-print-block:not(.dve-print-table)')
				.filter({ hasText: 'T2R0C0L00' })
				.first();

			await outer.click({ position: { x: 12, y: 8 } });
			await page.keyboard.press('Home');
			await page.keyboard.type('Ed');
			const sourceOuter = input.locator('p').filter({ hasText: 'T2R0C0L00' });
			await expect(sourceOuter).toHaveText('EdT2R0C0L00');
			await expect(outer).toHaveText('EdT2R0C0L00');
			await page.keyboard.press('Control+z');
			await expect(sourceOuter).toHaveText('T2R0C0L00');
			await expect(outer).toHaveText('T2R0C0L00');
			await page.keyboard.press('Control+y');
			await expect(sourceOuter).toHaveText('EdT2R0C0L00');
			await expect(outer).toHaveText('EdT2R0C0L00');

			const before = await input.textContent();
			await page.keyboard.press('End');
			await page.keyboard.press('Enter');
			await expect.poll(() => input.textContent()).toBe(before);
			const bytes = await save(),
				edited = await JSZip.loadAsync(bytes);
			expect(nested(await edited.file('word/document.xml')!.async('string'))).toBe(
				nested(await zip.file('word/document.xml')!.async('string')),
			);
			for (const [path, file] of Object.entries(zip.files))
				if (!file.dir && path !== 'word/document.xml')
					expect(await edited.file(path)!.async('uint8array')).toEqual(
						await file.async('uint8array'),
					);
			const oldSourceEditor = await input.elementHandle();
			const oldPrintPage = await editor.locator('.dve-print-page').last().elementHandle();
			await (
				await fileInput(page)
			).setInputFiles({
				name: 'reopened.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer: bytes,
			});

			await expect.poll(() => oldSourceEditor!.evaluate((node) => node.isConnected)).toBe(false);
			await expect.poll(() => oldPrintPage!.evaluate((node) => node.isConnected)).toBe(false);
			await oldSourceEditor!.dispose();
			await oldPrintPage!.dispose();
			await expect(input.locator('p').filter({ hasText: 'T2R0C0L00' })).toHaveText('EdT2R0C0L00');
			await print();
			await expect(outer).toHaveText('EdT2R0C0L00');

			await check();
		},
	);
