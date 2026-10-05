import { defineConfig } from '@playwright/test';
import { resolve } from 'node:path';
export default defineConfig({
	testDir: './tests',
	fullyParallel: true,
	use: {
		baseURL: 'http://127.0.0.1:4173',
		launchOptions: {
			...(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {}),
		},
	},
	webServer: {
		// The docs site and every demo are served from "/" out of one folder, as the Pages site is
		// served from its repository path.
		command:
			'node scripts/build-pages.mjs .browser-test-dist --browser-tests && node docs/node_modules/vitepress/bin/vitepress.js preview docs --host 127.0.0.1 --port 4173',
		env: { DOCS_BASE: '/', DOCS_OUT: resolve('.browser-test-dist') },
		url: 'http://127.0.0.1:4173/demo/',
		reuseExistingServer: false,
	},
});
