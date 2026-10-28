import { chromium, defineConfig } from '@playwright/test';
import { resolveChromiumExecutable } from '../../scripts/playwright-chromium.mjs';
import { browserPerformance } from '../../scripts/browser-performance.mjs';

const port = Number(process.env.PLAYWRIGHT_PORT ?? 4180);
const origin = `http://127.0.0.1:${port}`;
const executablePath = resolveChromiumExecutable(chromium.executablePath());

export default defineConfig({
	testDir: '../../e2e/docx',
	fullyParallel: false,
	// CI's runners have 4 vCPUs: two workers about halve the suite. Its timing-sensitive specs wait
	// for the editor to settle, so they stay green when the CPU is shared.
	workers: browserPerformance().workers,
	use: {
		baseURL: origin,
		headless: true,
		...(executablePath ? { launchOptions: { executablePath } } : {}),
	},
	webServer: {
		// Test the production build: the dev server issues hundreds of unbundled module requests per
		// page, which exhausts ephemeral ports on Windows (net::ERR_ADDRESS_IN_USE) over a full run.
		command: `bun run build && bun x vite preview ../../demos/docx/demo-vanilla --config vite.config.ts --host 127.0.0.1 --port ${port} --strictPort`,
		url: origin,
		reuseExistingServer: !process.env.CI,
		timeout: 180_000,
	},
	reporter: 'list',
});
