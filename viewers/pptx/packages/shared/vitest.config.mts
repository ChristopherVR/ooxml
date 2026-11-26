import { testWorkers } from '../../../../scripts/test-workers.mjs';

export default {
	test: {
		maxWorkers: testWorkers(),
		testTimeout: 30_000,
	},
};
