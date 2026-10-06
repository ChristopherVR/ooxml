import { defineConfig } from '@playwright/test';

const port = Number(process.env.PLAYWRIGHT_PORT ?? 4190);
const origin = `http://127.0.0.1:${port}`;

// The specs live in e2e/teams at the repository root and run against the vanilla demo
// (demos/teams/vanilla), on its own port, so a dev server already running on 5173 is never mistaken for it.
export default defineConfig({
	testDir: '../../e2e/teams',
	fullyParallel: false,
	workers: 1,
	use: { baseURL: origin, headless: true },
	webServer: {
		command: `bun x vite --host 127.0.0.1 --port ${port} --strictPort`,
		cwd: '../../demos/teams/vanilla',
		url: origin,
		reuseExistingServer: false,
		timeout: 120_000,
	},
	reporter: 'list',
});
