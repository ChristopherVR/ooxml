import { availableParallelism } from 'node:os';

/** Worker budget for one Vitest invocation. Vitest gives --maxWorkers precedence over config. */
export function testWorkers(env = process.env, cpus = availableParallelism()) {
	const requested = env.OOXML_TEST_WORKERS;
	if (requested) {
		const workers = Number(requested);
		if (!/^\d+$/.test(requested) || !Number.isSafeInteger(workers) || workers < 1) {
			throw new Error('OOXML_TEST_WORKERS must be a positive integer');
		}
		return workers;
	}
	return env.CI ? Math.min(4, cpus) : Math.max(1, Math.min(8, Math.floor(cpus / 2)));
}
