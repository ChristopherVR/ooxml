import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileInput } from './helpers';
const evidence = JSON.parse(
	readFileSync(
		new URL(
			'../../src/core/docx/__fixtures__/long-inner-notes/native-measurements.json',
			import.meta.url,
		),
		'utf8',
	),
) as {
	cases: {
		name: string;
		measurements: {
			lines: { text: string }[];
			rules: number[][];
			borders: { role: string; rectPt: number[] }[];
		}[];
	}[];
};
for (const sample of evidence.cases)
	test(`long interacting notes preserve continuation and source: ${sample.name}`, async ({
		page,
	}) => {
		await page.goto('/?framework=vanilla');
		const editor = page.locator('docx-editor'),
			input = editor.locator('.ProseMirror').first();
		const upload = async (buffer: Buffer) =>
			(await fileInput(page)).setInputFiles({
				name: sample.name + '.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer,
			});
		await upload(
			readFileSync(
				new URL(
					'../../src/core/docx/__fixtures__/long-inner-notes/' + sample.name + '.docx',
					import.meta.url,
				),
			),
		);
		await expect(input).toContainText('Outer cell');
		const print = async () => {
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			await editor
				.getByRole('combobox', { name: 'Layout view', exact: true })
				.selectOption('print');
		};
		await print();
		const check = async () => {
			const pages = editor.locator('.dve-print-page');
			await expect(pages).toHaveCount(sample.measurements.length);
			for (const [index, native] of sample.measurements.entries()) {
				const text = await pages
					.nth(index)
					.locator('.dve-print-block:not(.dve-print-table)')
					.allTextContents();
				expect(text.flatMap((text) => text.match(/T\dR\d+C\d+L\d{2}/g) ?? []).sort()).toEqual(
					native.lines.flatMap((line) => line.text.match(/T\dR\d+C\d+L\d{2}/g) ?? []).sort(),
				);
				const notes = await pages.nth(index).locator('.dve-print-footnote').allTextContents();
				expect(notes.flatMap((text) => text.match(/N\d+L\d{2}/g) ?? []).sort()).toEqual(
					native.lines.flatMap((line) => line.text.match(/N\d+L\d{2}/g) ?? []).sort(),
				);
				const tables = pages.nth(index).locator('.dve-print-table:not(.dve-print-nested-table)');
				const borders = native.borders.filter((border) => border.role === 'outer');
				if (!borders.length) await expect(tables).toHaveCount(0);
				else {
					await expect(tables).toHaveCount(1);
					const band = await tables
						.first()
						.evaluate((node) => parseFloat((node as HTMLElement).style.height));
					expect(
						Math.abs(
							band * 0.75 -
								(Math.max(...borders.map((border) => border.rectPt[3]!)) -
									Math.min(...borders.map((border) => border.rectPt[1]!))),
						),
					).toBeLessThan(0.2);
				}
				const noteLines = native.lines
					.map((line) => line.text.match(/N(\d+)L(\d{2})/))
					.filter((match) => match !== null);
				const firstById = new Map<string, number>();
				for (const match of noteLines)
					if (!firstById.has(match![1]!)) firstById.set(match![1]!, Number(match![2]));
				const continued = [...firstById.values()].filter((first) => first > 0).length;
				await expect(
					pages.nth(index).locator('.dve-print-footnote[data-continued="true"]'),
				).toHaveCount(continued);
				const rules = pages.nth(index).locator('.dve-print-note-rule');
				await expect(rules).toHaveCount(native.rules.length);
				for (const [ruleIndex, rule] of native.rules.entries()) {
					const width = await rules
						.nth(ruleIndex)
						.evaluate((node) => parseFloat((node as HTMLElement).style.width));
					expect(Math.abs(width * 0.75 - (rule[2]! - rule[0]!))).toBeLessThan(0.2);
				}
			}
		};
		await check();
		const direct = editor
			.locator('.dve-print-table:not(.dve-print-nested-table) > .dve-print-row > .dve-print-cell')
			.nth(1);
		await direct
			.locator('.dve-print-block')
			.first()
			.click({ position: { x: 15, y: 8 } });
		// Pixel hit testing can land inside the marker on another platform.
		await page.keyboard.press('Home');
		await page.keyboard.type('Edit');
		await expect(input).toContainText('EditT0R0C1L00');
		await expect(editor.locator('.dve-print-pages').first()).toContainText('EditT0R0C1L00');
		await page.keyboard.press('Control+z');
		await expect(input).not.toContainText('Edit');
		await expect(editor.locator('.dve-print-pages').first()).not.toContainText('Edit');
		await check();
		await page.keyboard.press('Control+y');
		await expect(input).toContainText('EditT0R0C1L00');
		await expect(editor.locator('.dve-print-pages').first()).toContainText('EditT0R0C1L00');
		const bytes = await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		const oldInput = await input.elementHandle();
		const oldPrintPage = await editor.locator('.dve-print-page').first().elementHandle();
		expect(oldInput).not.toBeNull();
		expect(oldPrintPage).not.toBeNull();
		await upload(Buffer.from(bytes));
		// The previous editor already contains Edit: wait for the actual reopened document.
		await expect.poll(() => oldInput!.evaluate((node) => node.isConnected)).toBe(false);
		await oldInput!.dispose();
		await expect(input).toContainText('EditT0R0C1L00');
		await print();
		// Reopening schedules a fresh print projection; identical text is not a generation signal.
		await expect.poll(() => oldPrintPage!.evaluate((node) => node.isConnected)).toBe(false);
		await oldPrintPage!.dispose();
		await expect(editor.locator('.dve-print-pages').first()).toContainText('EditT0R0C1L00');
		await check();
	});
