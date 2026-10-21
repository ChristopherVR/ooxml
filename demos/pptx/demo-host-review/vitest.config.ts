import { defineConfig } from 'vitest/config';

// The server's pure pieces run in Node; the page is checked in a browser (see README).
export default defineConfig({
	test: { include: ['server/**/*.test.ts', 'src/**/*.test.ts'], environment: 'node' },
});
