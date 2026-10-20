import { defineConfig } from '@playwright/test';
const port = Number(process.env.PLAYWRIGHT_PORT ?? 14450);
export default defineConfig({
	testDir: '../../e2e/visio',
	testMatch: ['native-building-blocks.spec.ts', 'native-theme.spec.ts', 'shadow-theme.spec.ts'],
	workers: 2,
	use: {
		baseURL: `http://127.0.0.1:${port}`,
		launchOptions: process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {},
	},
	webServer: {
		command: `bun x vite --config native-blocks.vite.config.ts --host 127.0.0.1 --port ${port}`,
		url: `http://127.0.0.1:${port}/native-blocks.html`,
		reuseExistingServer: false,
	},
});
