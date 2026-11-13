/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Host-owned Set Up Slide Show, Custom Shows, Broadcast and shortcut help of
 * `useViewerBuildingBlocks`, driven through the React demo's custom shell (`HostOwnedShowDialogs`).
 * The building blocks are a React-only API, so the other bindings skip.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { hostOwnedSessionUrl } from './support/host-owned-session';

test.beforeEach(() => {
	test.skip(test.info().project.name !== 'react', 'React-only building-block API');
});

const out = (page: Page, label: string) => page.getByLabel(label, { exact: true });
const click = (page: Page, name: string) => page.getByRole('button', { name, exact: true }).click();

async function open(page: Page, room: string, extra: Record<string, string> = {}): Promise<void> {
	await page.goto(hostOwnedSessionUrl(room, { headless: '1', sample: '1', ...extra }));
	await expect(page.locator('[data-host-show-dialogs]')).toBeVisible({ timeout: 30_000 });
	await expect(page.locator('[data-pptx-viewport] [data-element-id]').first()).toBeVisible({
		timeout: 30_000,
	});
	if (extra.editable !== '0') {
		await expect(
			page.getByRole('status', { name: 'Collaboration status', exact: true }),
		).toContainText('Editable', { timeout: 30_000 });
	}
}

test.describe('host-owned slide show dialogs', () => {
	test('set up show: request, current values, apply, validation code', async ({ page }) => {
		await open(page, `show-setup-${Date.now()}`);
		await click(page, 'Host set up show');
		await expect(out(page, 'Show dialog request')).toHaveText('setUpSlideShow');
		await expect(out(page, 'Show type')).toHaveText('presented');
		await click(page, 'Apply kiosk');
		await expect(out(page, 'Show dialog result')).toHaveText('setup: ok');
		await expect(out(page, 'Show dialog request')).toHaveText('none');
		await click(page, 'Host set up show');
		await expect(out(page, 'Show type')).toHaveText('kiosk');
		await expect(out(page, 'Show loop')).toHaveText('true');
		await out(page, 'Show range end').fill('999');
		await click(page, 'Apply range');
		await expect(out(page, 'Show dialog result')).toHaveText('setup: invalid-range');
		await out(page, 'Show range end').fill('2');
		await click(page, 'Apply range');
		await expect(out(page, 'Show dialog result')).toHaveText('setup: ok');
	});

	test('custom shows: create, rename, edit slides, start', async ({ page }) => {
		await open(page, `show-custom-${Date.now()}`);
		await click(page, 'Host custom shows');
		await expect(out(page, 'Show dialog request')).toHaveText('customShows');
		await click(page, 'Create custom show');
		await expect(out(page, 'Show dialog result')).toHaveText('create: ok');
		await expect(out(page, 'Custom shows list')).toHaveText('Highlights[2]');
		await click(page, 'Rename first show');
		await expect(out(page, 'Custom shows list')).toHaveText('Highlights 2[2]');
		await click(page, 'Keep one slide');
		await expect(out(page, 'Custom shows list')).toHaveText('Highlights 2[1]');
		await click(page, 'Start first show');
		await expect(out(page, 'Show dialog result')).toHaveText('start: ok');
	});

	test('custom shows: delete', async ({ page }) => {
		await open(page, `show-delete-${Date.now()}`);
		await click(page, 'Host custom shows');
		await click(page, 'Create custom show');
		await expect(out(page, 'Custom shows list')).toHaveText('Highlights[2]');
		await click(page, 'Delete first show');
		await expect(out(page, 'Custom shows list')).toHaveText('none');
	});

	test('broadcast is reported unsupported and raises nothing', async ({ page }) => {
		await open(page, `show-broadcast-${Date.now()}`);
		await expect(out(page, 'Broadcast support')).toHaveText('unsupported (collaboration-service)');
		await click(page, 'Host broadcast');
		await expect(out(page, 'Show dialog request')).toHaveText('none');
		await expect(out(page, 'Show dialog availability')).toContainText('broadcast=unsupported');
	});

	test('shortcut help toggles and lists the keymap as data', async ({ page }) => {
		await open(page, `show-shortcuts-${Date.now()}`);
		await expect(out(page, 'Shortcut help')).toHaveText('closed');
		await click(page, 'Host shortcuts');
		await expect(out(page, 'Shortcut help')).toContainText('open (');
		await expect(out(page, 'Show dialog request')).toHaveText('shortcuts');
		await expect(page.getByLabel('Host shortcut list').getByRole('listitem').first()).toContainText(
			'Ctrl/Cmd',
		);
		await click(page, 'Host shortcuts');
		await expect(out(page, 'Shortcut help')).toHaveText('closed');
		await expect(out(page, 'Show dialog request')).toHaveText('none');
	});

	test('a read-only host sees set up and custom shows refused', async ({ page }) => {
		await open(page, `show-readonly-${Date.now()}`, { editable: '0' });
		await expect(out(page, 'Show dialog availability')).toContainText('setUpSlideShow=read-only');
		await click(page, 'Host set up show');
		await expect(out(page, 'Show dialog request')).toHaveText('none');
	});
});
