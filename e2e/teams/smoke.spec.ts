import { expect, test } from '@playwright/test';

// Starter spec for the OpenTeams browser tests. It runs against the vanilla demo in local mode
// (?local=1: same-browser chat over BroadcastChannel, no server), so it needs nothing but the demo.
// Add specs beside it; the structure mirrors e2e/docx, e2e/xlsx and e2e/visio.

test('the vanilla demo mounts <teams-app> in local mode', async ({ page }) => {
	await page.goto('/?local=1&name=Ada');
	await expect(page.locator('teams-app')).toBeVisible();
	await expect(page).toHaveTitle(/OpenTeams/);
});
