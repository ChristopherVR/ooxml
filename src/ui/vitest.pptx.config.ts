import { fileURLToPath } from 'node:url';

const coreSource = fileURLToPath(new URL('../core', import.meta.url)).replaceAll('\\', '/');

export default {
	resolve: {
		alias: [{ find: /^ooxml-core\/(.+)$/, replacement: `${coreSource}/$1` }],
	},
	test: {
		css: { include: [/\.css\?raw$/] },
		globals: true,
		environment: 'node',
		include: ['src/pptx/**/*.test.ts'],
		maxWorkers: 4,
		testTimeout: 30000,
		hookTimeout: 30000,
	},
};
