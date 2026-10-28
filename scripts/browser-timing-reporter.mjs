import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';

/** One process writes each report. No shared file is mutated by workers or shards. */
export default class BrowserTimingReporter {
	tests = {};
	onBegin(config) {
		this.root = config.rootDir;
	}
	onTestEnd(test, result) {
		if (result.status !== 'passed' || test.expectedStatus !== 'passed') return;
		const key = JSON.stringify([
			test.parent.project().name,
			relative(this.root, test.location.file).replaceAll('\\', '/'),
			test.titlePath().slice(2),
		]);
		this.tests[key] = Math.max(1, result.duration);
	}
	onEnd() {
		const path = resolve(process.env.OOXML_E2E_TIMINGS_OUTPUT ?? 'test-results/timings.json');
		mkdirSync(dirname(path), { recursive: true });
		writeFileSync(path, JSON.stringify({ version: 1, tests: this.tests }, null, 2));
	}
}
