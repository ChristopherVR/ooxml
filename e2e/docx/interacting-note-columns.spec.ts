import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileInput } from './helpers';
const evidence = JSON.parse(
	readFileSync(
		new URL(
			'../../src/core/docx/__fixtures__/interacting-note-columns/native-measurements.json',
			import.meta.url,
		),
		'utf8',
	),
) as {
	cases: {
		name: string;
		measurements: {
			lines: { text: string; leftPt: number }[];
			borders: { role: string; rectPt: number[] }[];
		}[];
	}[];
};
for (const sample of evidence.cases)
	test(`interacting note columns retain native last-line completion: ${sample.name}`, async ({
		page,
	}) => {
		await page.goto('/?framework=vanilla');
		const editor = page.locator('docx-editor'),
			input = editor.locator('.ProseMirror').first();
		const upload = async (buffer: Buffer, name = sample.name + '.docx') =>
			(await fileInput(page)).setInputFiles({
				name,
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer,
			});
		await upload(
			readFileSync(
				new URL(
					'../../src/core/docx/__fixtures__/interacting-note-columns/' + sample.name + '.docx',
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
				const nativeOffsets = new Map<string, Set<number>>();
				const sourceFirst = new Map<string, number>();
				for (const line of native.lines) {
					const marker = line.text.match(/N(\d+)L(\d{2})/);
					if (!marker) continue;
					const key = String(Number(marker[1]) + 1),
						positions = nativeOffsets.get(key) ?? new Set<number>();
					positions.add(Math.round(line.leftPt));
					nativeOffsets.set(key, positions);
					sourceFirst.set(key, Math.min(sourceFirst.get(key) ?? Infinity, Number(marker[2])));
				}
				for (const [id, positions] of nativeOffsets) {
					const fragments = pages
						.nth(index)
						.locator('.dve-print-footnote[data-note-id="' + id + '"]');
					await expect(fragments).toHaveCount(positions.size);
					const actual = await fragments.evaluateAll((nodes) =>
						nodes
							.map((node) => Math.round(parseFloat((node as HTMLElement).style.left) * 0.75 + 40))
							.sort((a, b) => a - b),
					);
					expect(actual).toEqual([...positions].sort((a, b) => a - b));
					await expect(
						pages
							.nth(index)
							.locator('.dve-print-footnote[data-note-id="' + id + '"][data-continued="true"]'),
					).toHaveCount(sourceFirst.get(id)! > 0 ? positions.size : 0);
				}
			}
		};
		await check();
		const direct = editor
			.locator('.dve-print-table:not(.dve-print-nested-table) > .dve-print-row > .dve-print-cell')
			.nth(1);
		await direct.click({ position: { x: 15, y: 8 } });
		await page.keyboard.press('Home');
		await page.keyboard.type('Edit');
		await expect(input).toContainText('Edit');
		await page.keyboard.press('Control+z');
		await expect(input).not.toContainText('Edit');
		await check();
		await page.keyboard.press('Control+y');
		await expect(input).toContainText('Edit');
		const bytes = await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		const reopened = 'reopened-' + sample.name + '.docx';
		await upload(Buffer.from(bytes), reopened);
		await expect(editor.getByText(reopened, { exact: true })).toBeVisible();
		await expect(input).toContainText('Edit');
		await print();
		await check();
	});
