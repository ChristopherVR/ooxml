/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Collaboration through the building blocks of every binding that has them: two `?blocks=1`
 * pages join one host-owned room on the local relay through the blocks' `collaboration` option
 * (the shared host panel's "Host join room"), see each other in the shell state, share an edit,
 * and leave again without losing the open deck. One spec for all bindings with
 * `metadata.buildingBlocks`.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { requireBuildingBlocks } from './support/building-blocks';

test.beforeEach(() => requireBuildingBlocks());

const out = (page: Page, label: string) => page.getByLabel(label, { exact: true });
const press = (page: Page, name: string) => page.getByRole('button', { name, exact: true }).click();
const elementCount = (page: Page) =>
	page
		// Rendered elements only: a peer's remote-selection outline also carries `data-element-id`.
		.locator('[data-pptx-viewport] [data-pptx-element="true"][data-element-id]')
		.evaluateAll(
			(nodes) => new Set(nodes.map((node) => node.getAttribute('data-element-id') ?? '')).size,
		);

function roomUrl(room: string, intent: 'create' | 'join', name: string): string {
	const server =
		process.env.PPTX_E2E_COLLAB_SERVER ??
		`ws://127.0.0.1:${1234 + Number(process.env.PPTX_E2E_PORT_OFFSET ?? 0)}`;
	return `/?${new URLSearchParams({ blocks: '1', room, server, intent, name })}`;
}

async function join(page: Page, url: string): Promise<void> {
	await page.goto(url);
	await expect(page.locator('[data-building-blocks-panel]')).toBeVisible({ timeout: 30_000 });
	await expect(out(page, 'Blocks slides')).not.toHaveText('0', { timeout: 30_000 });
	await expect(out(page, 'Blocks collaboration')).toHaveText('off');
	await expect(out(page, 'Blocks connected')).toHaveText('0');
	await press(page, 'Grant editing');
	await expect(out(page, 'Blocks mode')).toHaveText('edit');
	await press(page, 'Host join room');
	await expect(out(page, 'Blocks collaboration')).toHaveText('connected', { timeout: 30_000 });
	await expect(out(page, 'Blocks shell editable')).toHaveText('true', { timeout: 30_000 });
}

test('two hosts share a room through the building blocks and leave it again', async ({
	browser,
	baseURL,
}, info) => {
	const room = `blocks-${info.project.name}-${Date.now()}`;
	const contexts = await Promise.all([0, 1].map(() => browser.newContext({ baseURL })));
	try {
		const [owner, peer] = await Promise.all(contexts.map((context) => context.newPage()));
		await join(owner!, roomUrl(room, 'create', 'Owner'));
		await join(peer!, roomUrl(room, 'join', 'Peer'));
		// Each sees the other in the shell state.
		await expect(out(owner!, 'Blocks connected')).toHaveText('2', { timeout: 30_000 });
		await expect(out(peer!, 'Blocks connected')).toHaveText('2', { timeout: 30_000 });

		// An insert on one reaches the other.
		const before = await elementCount(peer!);
		await press(owner!, 'Host SmartArt');
		await expect(out(owner!, 'Blocks result')).toHaveText('ok');
		await expect.poll(() => elementCount(peer!), { timeout: 30_000 }).toBe(before + 1);

		// Leaving: off again, still editable, the deck (with the shared edit) kept.
		await press(peer!, 'Host leave room');
		await expect(out(peer!, 'Blocks collaboration')).toHaveText('off', { timeout: 30_000 });
		await expect(out(peer!, 'Blocks connected')).toHaveText('0');
		await expect(out(peer!, 'Blocks mode')).toHaveText('edit');
		await expect(out(peer!, 'Blocks shell editable')).toHaveText('true');
		expect(await elementCount(peer!)).toBe(before + 1);
		await expect(out(owner!, 'Blocks connected')).toHaveText('1', { timeout: 30_000 });
	} finally {
		await Promise.all(contexts.map((context) => context.close()));
	}
});
