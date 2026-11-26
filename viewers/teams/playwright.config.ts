import { defineConfig } from '@playwright/test';

const port = Number(process.env.PLAYWRIGHT_PORT ?? 4190);
const origin = `http://127.0.0.1:${port}`;
const framework = process.env.TEAMS_TEST_FRAMEWORK ?? 'vanilla';
if (!['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'].includes(framework)) {
	throw new Error(`Unsupported OpenTeams test framework: ${framework}`);
}

// The specs live in e2e/teams at the repository root and run against the selected demo
// (demos/teams/vanilla), on its own port, so a dev server already running on 5173 is never mistaken for it.
export default defineConfig({
	testDir: '../../e2e/teams',
	fullyParallel: false,
	workers: 1,
	use: {
		baseURL: origin,
		headless: true,
		// `vite preview` gzips every response on the fly (about 0.2 s for a 10 MB bundle), and each
		// test's fresh context downloads the whole demo again; on localhost that only burns the
		// runner's CPU, so ask for uncompressed responses.
		extraHTTPHeaders: { 'Accept-Encoding': 'identity' },
	},
	webServer: {
		command: `bun x vite build && bun x vite preview --host 127.0.0.1 --port ${port} --strictPort`,
		cwd: `../../demos/teams/${framework}`,
		url: origin,
		reuseExistingServer: false,
		timeout: 120_000,
	},
	reporter: 'list',
});
