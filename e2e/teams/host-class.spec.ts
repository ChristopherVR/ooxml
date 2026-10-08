import { expect, test } from '@playwright/test';

// Every demo passes the same host class to its binding (React `className`, Vue `class`
// fallthrough, Solid and Svelte `class`, Angular `className`, vanilla `mountTeams({ className })`),
// and `window.teamsDemo.setHostClass` (demos/teams/test-hooks.ts) re-renders it with another.
type DemoWindow = Window & { teamsDemo: { setHostClass(value: string): void } };

test('the binding puts the host class on <teams-app> and swaps it on re-render', async ({
	page,
}) => {
	await page.goto(`/?local=1&name=Ada&room=host-class-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	const host = page.locator('teams-app');
	await expect(host).toHaveCount(1);
	await expect(host).toHaveClass(/(^|\s)demo-teams-host(\s|$)/);
	await page.evaluate(() =>
		(window as unknown as DemoWindow).teamsDemo.setHostClass('demo-teams-swapped extra'),
	);
	await expect(host).toHaveClass(/(^|\s)demo-teams-swapped(\s|$)/);
	await expect(host).toHaveClass(/(^|\s)extra(\s|$)/);
	await expect(host).not.toHaveClass(/(^|\s)demo-teams-host(\s|$)/);
	await page.evaluate(() => (window as unknown as DemoWindow).teamsDemo.setHostClass('back'));
	await expect(host).toHaveClass(/(^|\s)back(\s|$)/);
	await expect(host).not.toHaveClass(/demo-teams-swapped|extra/);
	// The element keeps working after the class changes.
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible();
});
