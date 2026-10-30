import { expect, test } from '@playwright/test';
import type { Workbook } from 'ooxml-core/xlsx';
import {
	editor,
	FRAMEWORKS,
	goToCell,
	grid,
	nameBoxValue,
	newWorkbook,
	openSample,
	pageErrors,
} from './helpers';

interface SelectionEditor extends HTMLElement {
	workbook: Workbook;
	getSelection(): string;
	setActiveSheet(index: number): void;
	saveBytes(): Promise<Uint8Array>;
	load(bytes: Uint8Array, fileName: string): Promise<void>;
}

for (const framework of FRAMEWORKS) {
	test(`${framework}: navigation and selection survive save and reload on every sheet`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await openSample(page, framework);
		await goToCell(page, 'B3');
		await grid(page).focus();
		await page.keyboard.press('Shift+ArrowRight');
		await expect
			.poll(() => editor(page).evaluate((node) => (node as SelectionEditor).getSelection()))
			.toBe('B3:C3');
		await editor(page).evaluate((node) => (node as SelectionEditor).setActiveSheet(1));
		await goToCell(page, 'D6');
		await grid(page).focus();
		await page.keyboard.press('ArrowRight');
		await expect.poll(() => nameBoxValue(page)).toBe('E6');
		await editor(page).evaluate(async (node) => {
			const el = node as SelectionEditor;
			await el.load(await el.saveBytes(), 'selection.xlsx');
		});
		await expect.poll(() => nameBoxValue(page)).toBe('E6');
		expect(
			await editor(page).evaluate((node) => (node as SelectionEditor).workbook.activeSheet),
		).toBe(1);
		await editor(page).evaluate((node) => (node as SelectionEditor).setActiveSheet(0));
		await expect
			.poll(() => editor(page).evaluate((node) => (node as SelectionEditor).getSelection()))
			.toBe('B3:C3');
		expect(errors).toEqual([]);
	});
	for (const kind of ['freeze', 'split'] as const) {
		test(`${framework}: pointer selection in ${kind} panes survives save`, async ({ page }) => {
			const errors = pageErrors(page);
			await newWorkbook(page, framework);
			await editor(page).evaluate(async (node, paneKind) => {
				const el = node as SelectionEditor;
				const view = el.workbook.sheets[0]!.view;
				if (paneKind === 'freeze') view.freeze = { rows: 4, cols: 3 };
				else view.split = { x: 3000, y: 1400 };
				await el.load(await el.saveBytes(), 'panes.xlsx');
			}, kind);
			await editor(page)
				.locator('.xg-q-main')
				.click({ position: { x: 100, y: 80 } });
			const mainSelection = await editor(page).evaluate(
				(node) => (node as SelectionEditor).workbook.sheets[0]!.view.selection,
			);
			await editor(page)
				.locator('.xg-q-corner')
				.click({ position: { x: 20, y: 15 } });
			const cornerSelection = await nameBoxValue(page);
			await editor(page).evaluate(async (node) => {
				const el = node as SelectionEditor;
				await el.load(await el.saveBytes(), 'panes.xlsx');
			});
			await expect.poll(() => nameBoxValue(page)).toBe(cornerSelection);
			const savedSplit = await editor(page).evaluate(
				(node) => (node as SelectionEditor).workbook.sheets[0]!.view.split,
			);
			if (kind === 'split') {
				expect(savedSplit?.activePane).toBe('topLeft');
				expect(savedSplit?.paneSelections?.bottomRight).toEqual(mainSelection);
			} else expect(savedSplit).toBeUndefined();
			expect(errors).toEqual([]);
		});
	}
}
