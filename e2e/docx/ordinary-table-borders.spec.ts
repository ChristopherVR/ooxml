import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileInput } from './helpers';
const native = JSON.parse(
	readFileSync(
		new URL(
			'../../src/core/docx/__fixtures__/ordinary-table-borders/native-measurements.json',
			import.meta.url,
		),
		'utf8',
	),
) as {
	cases: { name: string; measurements: { borders: { role: string; rectPt: number[] }[] }[] }[];
};
for (const name of [
	'ordinary-border-short',
	'ordinary-border-half',
	'ordinary-border-margin',
	'ordinary-border-mixed',
	'enclosing-border-auto',
	'enclosing-border-exact',
	'enclosing-border-clipped',
	'ordinary-border-split-early',
	'ordinary-border-split-late',
	'ordinary-border-header',
]) {
	test(`ordinary bordered rows preserve physical geometry and source: ${name}`, async ({
		page,
	}) => {
		await page.goto('/?framework=vanilla');
		const editor = page.locator('docx-editor');
		const input = editor.locator('.ProseMirror').first();
		const upload = async (buffer: Buffer) =>
			(await fileInput(page)).setInputFiles({
				name: name + '.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer,
			});
		await upload(
			readFileSync(
				new URL(
					'../../src/core/docx/__fixtures__/ordinary-table-borders/' + name + '.docx',
					import.meta.url,
				),
			),
		);
		await expect(input).toContainText(name.startsWith('enclosing') ? 'Outer cell' : 'T0');
		const print = async () => {
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			await editor
				.getByRole('combobox', { name: 'Layout view', exact: true })
				.selectOption('print');
			await expect(editor.locator('.dve-print-pages').first()).toBeVisible();
		};
		await print();
		const cell = editor
			.locator('.dve-print-table:not(.dve-print-nested-table) > .dve-print-row > .dve-print-cell')
			.first();
		await expect(cell).toHaveCSS('border-top-color', 'rgb(255, 0, 0)');
		const snapshot = async () =>
			editor.locator('.dve-print-table:not(.dve-print-nested-table)').evaluateAll((nodes) =>
				nodes.map((node) => ({
					height: node.getBoundingClientRect().height,
					rows: Array.from(node.querySelectorAll(':scope > .dve-print-row')).map(
						(row) => row.getBoundingClientRect().height,
					),
				})),
			);
		const before = await snapshot();
		const sample = native.cases.find((sample) => sample.name === name)!;
		const sourceHeights = await editor
			.locator('.dve-print-table:not(.dve-print-nested-table)')
			.evaluateAll((nodes) => nodes.map((node) => parseFloat((node as HTMLElement).style.height)));
		for (const [index, height] of sourceHeights.entries()) {
			const borders = sample.measurements[index]!.borders.filter(
				(border) => border.role === 'outer',
			);
			const top = Math.min(...borders.map((border) => border.rectPt[1]!));
			const bottom = Math.max(...borders.map((border) => border.rectPt[3]!));
			expect(Math.abs(height * 0.75 - (bottom - top))).toBeLessThan(0.2);
		}
		const pages = editor.locator('.dve-print-page');
		if (name === 'ordinary-border-split-early') {
			await expect(pages.nth(0).locator('.dve-print-footnote')).toContainText('N0L02');
			await expect(pages.nth(0).locator('.dve-print-cell').first()).toContainText('T0R0C0L12');
			await expect(pages.nth(0).locator('.dve-print-cell').first()).not.toContainText('T0R0C0L13');
		}
		if (name === 'ordinary-border-split-late') {
			await expect(pages.nth(0).locator('.dve-print-footnote')).toHaveCount(0);
			await expect(pages.nth(1).locator('.dve-print-footnote')).toContainText('N0L02');
			await expect(pages.nth(0).locator('.dve-print-cell').nth(1)).toContainText('T0R0C1L16');
		}
		if (name === 'ordinary-border-header')
			await expect(editor.locator('.dve-print-row[data-repeated="true"]')).toHaveCount(2);
		expect(before.length).toBe(
			name === 'ordinary-border-header' ? 3 : name.includes('split') ? 2 : 1,
		);
		if (name.startsWith('enclosing')) {
			const trailing = editor
				.locator('.dve-print-block:not(.dve-print-table)')
				.filter({ hasText: 'Outer cell' })
				.first();
			const bounds = await trailing.evaluate((node) => ({
				ink: node.getBoundingClientRect().toJSON(),
				clip: node.parentElement!.getBoundingClientRect().toJSON(),
			}));
			if (name === 'enclosing-border-clipped')
				expect(bounds.ink.bottom).toBeGreaterThan(bounds.clip.bottom);
			else expect(bounds.ink.bottom).toBeLessThanOrEqual(bounds.clip.bottom + 0.5);
		}
		if (!name.startsWith('enclosing')) {
			await cell.click({ position: { x: 15, y: 8 } });
			await page.keyboard.type('Edit');
			await expect(input).toContainText('Edit');
			await page.keyboard.press('Control+z');
			await expect(input).not.toContainText('Edit');
			await expect.poll(snapshot).toEqual(before);
		}
		const bytes = await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		await upload(Buffer.from(bytes));
		await expect(input).toContainText(name.startsWith('enclosing') ? 'Outer cell' : 'T0');
		await print();
		await expect.poll(snapshot).toEqual(before);
		await expect(cell).toHaveCSS('border-top-color', 'rgb(255, 0, 0)');
	});
}
