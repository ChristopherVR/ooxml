export default {
	test: {
		globals: true,
		environment: 'node',
		include: ['src/**/*.test.ts'],
		// The pptx area's integration tests load real multi-megabyte decks through the full
		// parse pipeline (and several round-trip them through save); the 5s default timed out on CI.
		testTimeout: 30_000,
		hookTimeout: 30_000,
		maxWorkers: 4,
	},
};
