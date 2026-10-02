import { defineConfig } from 'vitest/config';
import { svelte } from '@sveltejs/vite-plugin-svelte';
import solid from 'vite-plugin-solid';
export default defineConfig({
	plugins: [svelte(), solid({ hot: false })],
	resolve: { conditions: ['browser'] },
	test: {
		environment: 'jsdom',
		include: ['tests/**/*.test.ts'],
		exclude: ['tests/**/*.ssr.test.ts'],
		setupFiles: ['./tests/setup.ts'],
		isolate: true,
	},
});
