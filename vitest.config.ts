// The pptx tests read real decks from src/pptx/__tests__/fixtures (including the e2e snapshot that is
// committed under fixtures/e2e), so nothing outside this repository is needed.
export default {
	test: {
		globals: true,
		environment: 'node',
		include: ['src/**/*.test.ts'],
		// src/ui is its own package (ooxml-ui) with its own test setup.
		exclude: ['node_modules', 'dist', 'src/ui/**'],
		// The pptx area's integration tests load real multi-megabyte decks through the full
		// parse pipeline (and several round-trip them through save); the 5s default timed out on CI.
		testTimeout: 30_000,
		hookTimeout: 30_000,
		maxWorkers: 4,
	},
};
