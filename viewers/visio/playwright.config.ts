import { defineConfig } from '@playwright/test';
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
		command:
			'npm exec vite build -- --mode browser-tests && npm exec vite preview -- --outDir .browser-test-dist --host 127.0.0.1 --port 4173',
		url: 'http://127.0.0.1:4173/demo/',
		reuseExistingServer: false,
	},
});
