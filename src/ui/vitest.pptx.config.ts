export default {
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
