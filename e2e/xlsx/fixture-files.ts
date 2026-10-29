import { readFileSync } from 'node:fs';
import { expect, type Page } from '@playwright/test';
import { grid, openLanding } from './helpers';

const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** Reads a committed fixture of the core's `xlsx` area (Excel-made unless the spec says so). */
export const excelFixture = (name: string): Buffer =>
	readFileSync(new URL(`../../src/core/xlsx/__fixtures__/${name}`, import.meta.url));

/** Opens the demo in `framework` and loads a fixture through the landing file input. */
export async function openFixture(page: Page, framework: string, name: string) {
	await openLanding(page, framework);
	await page
		.locator('#landing-file')
		.setInputFiles({ name, mimeType: MIME, buffer: excelFixture(name) });
	await expect(grid(page)).toBeVisible();
}
