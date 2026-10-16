import { expect, test, type Page, type Route } from '@playwright/test';

// A stand-in for the reference server's HTTP side: /health reports accounts, /auth/session accepts
// one password. The sockets cannot connect (nothing listens), which the workspace tolerates.
const ORIGIN = 'http://127.0.0.1:8787';
const CORS = {
	'access-control-allow-origin': '*',
	'access-control-allow-headers': 'Content-Type, Authorization',
	'access-control-allow-methods': 'GET, POST, OPTIONS',
	'access-control-expose-headers': 'Retry-After',
};

async function fakeServer(page: Page, options: { accounts: boolean; limitAfter?: number }) {
	const state = { attempts: 0, signedOut: 0 };
	const reply = (
		route: Route,
		status: number,
		body?: unknown,
		headers: Record<string, string> = {},
	) =>
		route.fulfill({
			status,
			headers: { ...CORS, ...headers },
			...(body === undefined
				? {}
				: { contentType: 'application/json', body: JSON.stringify(body) }),
		});
	await page.route(`${ORIGIN}/**`, async (route) => {
		const request = route.request();
		const path = new URL(request.url()).pathname;
		if (request.method() === 'OPTIONS') return reply(route, 204);
		if (path === '/health') return reply(route, 200, { ok: true, accounts: options.accounts });
		if (path === '/auth/signout') {
			state.signedOut++;
			return reply(route, 204);
		}
		if (path === '/auth/session') {
			state.attempts++;
			if (options.limitAfter !== undefined && state.attempts > options.limitAfter)
				return reply(route, 429, undefined, { 'Retry-After': '30' });
			const body = request.postDataJSON() as { user: string; credential?: string };
			if (body.credential !== 'right-password') return reply(route, 401);
			return reply(route, 200, {
				session: 'a.b.c',
				user: body.user,
				expiresAt: Date.now() + 3_600_000,
			});
		}
		return reply(route, 404);
	});
	return state;
}

const useServerConfig = (page: Page) =>
	page.addInitScript(() => {
		localStorage.setItem(
			'teams:config',
			JSON.stringify({
				mode: 'server',
				syncUrl: 'ws://127.0.0.1:8787/sync',
				signalingUrl: 'ws://127.0.0.1:8787/signal',
				iceServers: [],
			}),
		);
	});

test.use({ viewport: { width: 1366, height: 768 } });

test('sign-in: wrong password, rate limit, sign in, profile menu, sign out', async ({ page }) => {
	await useServerConfig(page);
	const server = await fakeServer(page, { accounts: true, limitAfter: 1 });
	await page.goto('/?room=sign-in-spec');
	const form = page.locator('teams-sign-in');
	await expect(form.getByRole('heading', { name: 'Sign in to Teams' })).toBeVisible();

	await form.getByLabel('User ID').fill('alice');
	await form.getByLabel('Password').fill('wrong');
	await form.getByRole('button', { name: 'Sign in', exact: true }).click();
	await expect(form.getByRole('alert')).toHaveText('The user ID or password is not correct.');
	await expect(form.getByLabel('Password')).toHaveValue('');
	await expect(form.getByLabel('User ID')).toHaveValue('alice');

	// The next attempt reaches the rate limit: the wait is shown and the button is disabled.
	await form.getByLabel('Password').fill('right-password');
	await form.getByRole('button', { name: 'Sign in', exact: true }).click();
	await expect(form.getByRole('alert')).toContainText('Too many attempts');
	await expect(form.getByRole('button', { name: 'Sign in', exact: true })).toBeDisabled();
	expect(server.attempts).toBe(2);
});

test('sign-in: signing in opens the workspace as the account and sign out returns', async ({
	page,
}) => {
	await useServerConfig(page);
	const server = await fakeServer(page, { accounts: true });
	await page.goto('/?room=sign-in-ok');
	const form = page.locator('teams-sign-in');
	await form.getByLabel('User ID').fill('alice');
	await form.getByLabel('Password').fill('right-password');
	await form.getByRole('button', { name: 'Sign in', exact: true }).click();
	await expect(page.locator('teams-app .shell, teams-app').first()).toBeVisible();
	await expect(form).toHaveCount(0);
	await page.getByRole('button', { name: /^Profile for/u }).click();
	await expect(page.getByText('Signed in as alice')).toBeVisible();
	await page.getByRole('button', { name: 'Sign out', exact: true }).click();
	await expect(page.locator('teams-sign-in').getByText('You signed out.')).toBeVisible();
	await expect.poll(() => server.signedOut).toBe(1);
});

test('sign-in: a server without accounts never shows the form', async ({ page }) => {
	await useServerConfig(page);
	await fakeServer(page, { accounts: false });
	await page.goto('/?name=Ada&room=sign-in-dev');
	await expect(page.getByRole('button', { name: /^Profile for/u })).toBeVisible({
		timeout: 15_000,
	});
	await expect(page.locator('teams-sign-in')).toHaveCount(0);
});
