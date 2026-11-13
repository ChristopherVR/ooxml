/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * The building blocks of every binding that has them, through each demo's `?blocks=1` page: the
 * binding's own canvas, toolbar and slides pane plus the shared host panel
 * (`demos/pptx/shared/building-blocks-panel.ts`) that finishes their workflows through
 * `actions`. One spec for all of them; a binding joins by adding its demo page and
 * `metadata.buildingBlocks` to its Playwright project.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { requireBuildingBlocks } from './support/building-blocks';

test.beforeEach(() => requireBuildingBlocks());

const PNG = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
	'base64',
);

const out = (page: Page, label: string) => page.getByLabel(label, { exact: true });
const press = (page: Page, name: string) => page.getByRole('button', { name, exact: true }).click();
const elements = (page: Page) => page.locator('[data-pptx-viewport] [data-element-id]');
const elementIds = (page: Page) =>
	elements(page).evaluateAll((nodes) => [
		...new Set(nodes.map((node) => node.getAttribute('data-element-id') ?? '')),
	]);

async function open(page: Page, edit: boolean): Promise<void> {
	await page.goto('/?blocks=1');
	await expect(page.locator('[data-building-blocks-panel]')).toBeVisible({ timeout: 30_000 });
	await expect(elements(page).first()).toBeVisible({ timeout: 30_000 });
	await expect(out(page, 'Blocks slides')).not.toHaveText('0');
	if (edit) {
		await press(page, 'Grant editing');
		await expect(out(page, 'Blocks mode')).toHaveText('edit');
	}
}

test.describe('building blocks', () => {
	test('view-only until the host grants editing, then inserts run', async ({ page }) => {
		await open(page, false);
		await expect(out(page, 'Blocks mode')).toHaveText('preview');
		await press(page, 'Host SmartArt');
		await expect(out(page, 'Blocks result')).toHaveText('read-only');
		const before = (await elementIds(page)).length;
		await press(page, 'Grant editing');
		await expect(out(page, 'Blocks mode')).toHaveText('edit');
		await press(page, 'Host SmartArt');
		await expect(out(page, 'Blocks result')).toHaveText('ok');
		await expect.poll(async () => (await elementIds(page)).length).toBe(before + 1);
		await press(page, 'Revoke editing');
		await expect(out(page, 'Blocks mode')).toHaveText('preview');
	});

	test('a host-picked picture is inserted on the slide', async ({ page }) => {
		await open(page, true);
		const before = (await elementIds(page)).length;
		await out(page, 'Host picture file').setInputFiles({
			name: 'dot.png',
			mimeType: 'image/png',
			buffer: PNG,
		});
		await expect(out(page, 'Blocks result')).toHaveText('ok');
		await expect.poll(async () => (await elementIds(page)).length).toBe(before + 1);
	});

	test('Ctrl+K on a selected element raises the hyperlink request', async ({ page }) => {
		await open(page, true);
		await elements(page).filter({ hasText: /\w/u }).first().click();
		await page.keyboard.press('Control+k');
		await expect(out(page, 'Blocks pending')).toHaveText('hyperlink');
		await press(page, 'Host link');
		await expect(out(page, 'Blocks result')).toHaveText('ok');
		await expect(out(page, 'Blocks pending')).toHaveText('none');
	});

	test('the host find bar searches and steps through the matches', async ({ page }) => {
		await open(page, false);
		await out(page, 'Host find query').fill('e');
		await press(page, 'Host find');
		await expect
			.poll(async () => Number(await out(page, 'Blocks matches').textContent()))
			.toBeGreaterThan(1);
		await expect(out(page, 'Blocks current match')).toHaveText('0');
		await press(page, 'Host next match');
		await expect(out(page, 'Blocks current match')).toHaveText('1');
	});

	test('the dialog groups apply through actions', async ({ page }) => {
		await open(page, true);
		await press(page, 'Host set title');
		await expect(out(page, 'Blocks title')).toHaveText('Blocks title');
		await press(page, 'Host loop show');
		await expect(out(page, 'Blocks loop')).toHaveText('true');
		await press(page, 'Host add show');
		await expect(out(page, 'Blocks custom shows')).toHaveText('1');
		await press(page, 'Host shortcuts');
		await expect(out(page, 'Blocks shortcuts')).toHaveText('open');
		await expect(out(page, 'Blocks pending')).toHaveText('shortcuts');
		await press(page, 'Host cancel');
		await expect(out(page, 'Blocks shortcuts')).toHaveText('closed');
		await expect(out(page, 'Blocks pending')).toHaveText('none');
	});
});
