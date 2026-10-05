import { defineConfig } from 'vitest/config';

export default defineConfig({
	test: {
		environment: 'jsdom',
		include: ['src/**/*.test.ts'],
		// Component tests build a full chrome in jsdom; under the parallel full run (vitest 5, jsdom 30)
		// one took over 5 s. It passes alone in under 3 s, so this is load, not a hang.
		testTimeout: 20_000,
	},
});
