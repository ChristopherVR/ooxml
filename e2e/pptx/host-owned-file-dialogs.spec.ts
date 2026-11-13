/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Host-owned password protection, document properties, font embedding and digital signature
 * workflows of `useViewerBuildingBlocks` (#285), driven through the React demo's custom shell
 * (`HostOwnedFileDialogs`). The building blocks are a React-only API, so the other bindings skip.
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
	await expect(page.locator('[data-host-file-dialogs]')).toBeVisible({ timeout: 30_000 });
	await expect(page.locator('[data-pptx-viewport] [data-element-id]').first()).toBeVisible({
		timeout: 30_000,
	});
	if (extra.editable !== '0') {
		await expect(
			page.getByRole('status', { name: 'Collaboration status', exact: true }),
		).toContainText('Editable', { timeout: 30_000 });
	}
}

test.describe('host-owned file dialogs', () => {
	test('password protection: request, validation codes, state without secrets', async ({
		page,
	}) => {
		await open(page, `bb-protect-${Date.now()}`);
		await click(page, 'Host protection');
		await expect(out(page, 'Host pending request')).toHaveText('passwordProtection');
		await expect(out(page, 'Host protection state')).toContainText('open=false');
		await click(page, 'Set password');
		await expect(out(page, 'Host file dialog result')).toHaveText('password: password-required');
		await page.getByLabel('Host password', { exact: true }).fill('abc');
		await page.getByLabel('Host password confirmation').fill('abc');
		await click(page, 'Set password');
		await expect(out(page, 'Host file dialog result')).toHaveText('password: password-too-short');
		await page.getByLabel('Host password', { exact: true }).fill('secret1');
		await page.getByLabel('Host password confirmation').fill('secret2');
		await click(page, 'Set password');
		await expect(out(page, 'Host file dialog result')).toHaveText('password: password-mismatch');
		await page.getByLabel('Host password confirmation').fill('secret1');
		await click(page, 'Set password');
		await expect(out(page, 'Host file dialog result')).toHaveText('password: ok');
		await expect(out(page, 'Host pending request')).toHaveText('none');
		await click(page, 'Host protection');
		await expect(out(page, 'Host protection state')).toContainText('open=true');
		await expect(out(page, 'Host protection state')).not.toContainText('secret1');
		await click(page, 'Mark as final');
		await expect(out(page, 'Host protection state')).toContainText('final=true');
		await click(page, 'Host protection');
		await click(page, 'Set modify password');
		await expect(out(page, 'Host file dialog result')).toHaveText('modify: unsupported');
		await click(page, 'Clear password');
		await expect(out(page, 'Host protection state')).toContainText('open=false');
	});

	test('document properties: patch, malformed patch refused', async ({ page }) => {
		await open(page, `bb-props-${Date.now()}`);
		await click(page, 'Host properties');
		await expect(out(page, 'Host pending request')).toHaveText('documentProperties');
		await page.getByLabel('Host title').fill('Host title');
		await click(page, 'Apply bad property');
		await expect(out(page, 'Host file dialog result')).toHaveText('properties: invalid-data');
		await click(page, 'Apply properties');
		await expect(out(page, 'Host file dialog result')).toHaveText('properties: ok');
		await click(page, 'Host properties');
		await expect(out(page, 'Host properties state')).toContainText('title=Host title');
		await expect(out(page, 'Host properties state')).toContainText('Reviewer:Host');
	});

	test('font embedding: state, and embedding is refused when the deck carries no fonts', async ({
		page,
	}) => {
		await open(page, `bb-fonts-${Date.now()}`);
		await click(page, 'Host fonts');
		await expect(out(page, 'Host pending request')).toHaveText('fontEmbedding');
		await expect(out(page, 'Host fonts state')).toContainText('canEmbed=false');
		await click(page, 'Embed fonts');
		await expect(out(page, 'Host file dialog result')).toHaveText('fonts: unsupported');
		await click(page, 'Do not embed fonts');
		await expect(out(page, 'Host file dialog result')).toHaveText('fonts: ok');
	});

	test('digital signatures: unsigned deck lists none and creating one is unsupported', async ({
		page,
	}) => {
		await open(page, `bb-sigs-${Date.now()}`);
		await click(page, 'Host signatures');
		await expect(out(page, 'Host pending request')).toHaveText('digitalSignatures');
		await expect(out(page, 'Host signatures state')).toContainText('signed=false');
		await expect(out(page, 'Host signatures state')).toContainText('listed=0');
		await expect(out(page, 'Host signatures state')).toContainText('create=false');
		await click(page, 'Remove signatures');
		await expect(out(page, 'Host file dialog result')).toHaveText('signatures: no-signatures');
		await click(page, 'Sign');
		await expect(out(page, 'Host file dialog result')).toHaveText('signatures: unsupported');
	});

	test('read-only viewer shows the state and refuses every change', async ({ page }) => {
		await open(page, `bb-file-readonly-${Date.now()}`, { editable: '0' });
		await expect(out(page, 'Host action availability')).toContainText(
			'documentProperties=read-only*',
		);
		await click(page, 'Host properties');
		await expect(out(page, 'Host pending request')).toHaveText('documentProperties');
		await click(page, 'Apply properties');
		await expect(out(page, 'Host file dialog result')).toHaveText('properties: read-only');
		await click(page, 'Host protection');
		await click(page, 'Mark as final');
		await expect(out(page, 'Host file dialog result')).toHaveText('final: read-only');
		await expect(out(page, 'Host protection state')).toContainText('final=false');
	});
});
