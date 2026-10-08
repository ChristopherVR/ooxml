import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

// Every demo records what its binding's `open-file` handler receives in
// `window.teamsDemo.openedFiles` (demos/teams/test-hooks.ts): `(detail, event)` everywhere, and
// `{ detail, event }` unpacked by the Angular demo.
type DemoWindow = Window & {
	teamsDemo: {
		openedFiles: { name: string; kind: string; url: string | undefined; eventType?: string }[];
	};
};

test('opening a shared file from the Files view raises open-file naming the file', async ({
	page,
}) => {
	await page.route('**/content/Brief.md', (route) =>
		route.fulfill({ body: '# Launch brief\nShip it', contentType: 'text/markdown' }),
	);
	await page.goto(`/?local=1&name=Ada&room=open-file-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.locator('teams-app').evaluate((element) => {
		const client = (element as TeamsApp).client!;
		client.workspace.chat.post(client.getState().selectedChannelId, {
			text: 'Brief attached',
			attachments: [
				{ name: 'Brief.md', kind: 'other', url: `${location.origin}/content/Brief.md` },
			],
		});
	});
	const rail = page.getByRole('navigation', { name: 'App bar', exact: true });
	await rail.getByRole('button', { name: 'Files', exact: true }).click();
	await expect(page.getByRole('heading', { name: 'Files', exact: true })).toBeVisible();
	await page.getByRole('button', { name: 'Open Brief.md', exact: true }).click();
	await expect
		.poll(() => page.evaluate(() => (window as unknown as DemoWindow).teamsDemo.openedFiles))
		.toEqual([
			{
				name: 'Brief.md',
				kind: 'other',
				url: expect.stringContaining('/content/Brief.md'),
				eventType: 'teams-open-file',
			},
		]);
	// Not cancelled by the handler, so the workspace opens the file itself.
	await expect(page.getByRole('heading', { name: 'Launch brief', exact: true })).toBeVisible();
});
