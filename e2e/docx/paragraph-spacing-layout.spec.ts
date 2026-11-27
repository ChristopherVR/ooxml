import { expect, test, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { reopenDocument } from './helpers';

const fixtures = new URL(
	'../../src/core/docx/layout/fixtures/word-layout-reports/',
	import.meta.url,
);

async function open(page: Page, name: string, view: 'draft' | 'print') {
	await page.setViewportSize({ width: 1400, height: 900 });
	await page.goto('/?framework=vanilla');
	await reopenDocument(page, `${name}.docx`, await readFile(new URL(`${name}.docx`, fixtures)));
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	const layout = editor.getByRole('combobox', { name: 'Layout view', exact: true });
	await layout.selectOption(view);
	await expect(layout).toHaveValue(view);
	return editor;
}

test('the editing view keeps Word contextual spacing gaps (#359)', async ({ page }) => {
	const word = JSON.parse(
		await readFile(new URL('issue-359-contextual-spacing-rows.json', fixtures), 'utf8'),
	) as { gapsPt: Record<string, number> };
	const editor = await open(page, 'issue-359-contextual-spacing-rows', 'draft');
	await expect(editor.locator('.ProseMirror p').first()).toHaveText('row1S');
	const gaps = await editor.locator('.ProseMirror').evaluate((root, ids) => {
		const paragraphs = new Map(
			Array.from(root.querySelectorAll('p'), (p) => [p.textContent ?? '', p]),
		);
		return Object.fromEntries(
			ids.map((id) => {
				const upper = paragraphs.get(`${id}U`)?.getBoundingClientRect();
				const lower = paragraphs.get(`${id}L`)?.getBoundingClientRect();
				return [id, upper && lower ? (lower.top - upper.bottom) * 0.75 : NaN];
			}),
		);
	}, Object.keys(word.gapsPt));
	for (const [id, expected] of Object.entries(word.gapsPt))
		expect(Math.abs((gaps[id] ?? NaN) - expected), id).toBeLessThan(0.5);
});

test('Print Layout keeps the unflagged paragraph space after (#359)', async ({ page }) => {
	const editor = await open(page, 'issue-359-contextual-spacing-one-sided', 'print');
	await expect(editor.locator('.dve-print-page')).toHaveCount(2);
	const control = await open(page, 'issue-359-contextual-spacing-both-control', 'print');
	await expect(control.locator('.dve-print-page')).toHaveCount(1);
});

test('Print Layout puts multiple line spacing below the text (#360)', async ({ page }) => {
	const editor = await open(page, 'issue-360-line-spacing-rules', 'print');
	await expect(editor.locator('.dve-print-page')).toHaveCount(1);
	const tops = await editor.locator('.dve-print-page').evaluate((sheet) => {
		const result: Record<string, number> = {};
		const walker = document.createTreeWalker(sheet, NodeFilter.SHOW_TEXT);
		for (let node = walker.nextNode(); node; node = walker.nextNode()) {
			const text = node.textContent?.trim();
			if (!text || text in result || !node.parentElement) continue;
			result[text] = node.parentElement.getBoundingClientRect().top;
		}
		return result;
	});
	const step = (from: string, to: string) => ((tops[to] ?? NaN) - (tops[from] ?? NaN)) * 0.75;
	const single = step('single1', 'single2');
	// Each group follows a single line (the previous group's Sep line): its first line steps like
	// a single line, since the extra space goes below the text. The later lines step the multiple.
	const groups = [
		['singles', 'm115', 1.15],
		['m115s', 'm150', 1.5],
		['m150s', 'm200', 2],
	] as const;
	for (const [previous, group, multiple] of groups) {
		expect(Math.abs(step(previous, `${group}0`) - single), group).toBeLessThan(0.75);
		expect(Math.abs(step(`${group}0`, `${group}1`) - single * multiple), group).toBeLessThan(0.75);
	}
});
