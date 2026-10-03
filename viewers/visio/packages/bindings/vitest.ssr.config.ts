import { defineConfig } from 'vitest/config';
import { svelte } from '@sveltejs/vite-plugin-svelte';
import solid from 'vite-plugin-solid';
export default defineConfig({
	plugins: [svelte(), solid({ ssr: true })],
	resolve: { conditions: ['node'] },
	// The shared element source (and its ?inline CSS) lives in the repository root src/.
	server: { fs: { allow: ['../..'] } },
	test: { environment: 'node', include: ['tests/**/*.ssr.test.ts'], isolate: true },
});
