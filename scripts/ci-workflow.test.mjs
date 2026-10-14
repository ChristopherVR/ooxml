import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';

const workflow = readFileSync(new URL('../.github/workflows/ci.yml', import.meta.url), 'utf8');
const browserJob = workflow.match(/^  pptx-e2e:\r?\n([\s\S]*?)(?=^  \S)/m)?.[1];

test('incremental push plans cannot cancel checks for an earlier push', () => {
	const concurrency = workflow.match(/^concurrency:\r?\n([\s\S]*?)(?=^\S)/m)?.[1];
	assert.ok(concurrency);
	assert.match(
		concurrency,
		/group: ci-\$\{\{ github.event_name == 'pull_request' && github.event.pull_request.number \|\| github.run_id \}\}/u,
	);
	assert.match(concurrency, /cancel-in-progress: \$\{\{ github.event_name == 'pull_request' \}\}/u);
});

test('browser retries retain failed traces, including when the retry passes', () => {
	assert.ok(browserJob);
	const command = browserJob.match(/        run: \|\r?\n([\s\S]*?)(?=^      -)/m)?.[1];
	assert.ok(command, 'the browser job has a runner command');
	const script = command.replace(/^          /gm, '');
	for (const retryStatus of [0, 1]) {
		const dir = mkdtempSync(join(tmpdir(), 'ooxml-ci-retry-'));
		try {
			// Mimic Playwright clearing output before retrying. Use shell functions rather than
			// executables so this exercises the workflow on Windows and Linux alike.
			const harness = `
attempts=0
bun() {
  attempts=$((attempts + 1))
  printf '%s\\n' "$*" >> commands.txt
  rm -rf test-results playwright-report
  mkdir -p test-results playwright-report
  printf 'trace-%s' "$attempts" > test-results/trace.zip
  printf 'report-%s' "$attempts" > playwright-report/index.html
  printf '{"status":"failed","failedTests":["test-id"]}' > test-results/.last-run.json
  if [ "$attempts" -eq 1 ]; then return 1; fi
  return ${retryStatus}
}
sleep() { :; }
${script}`;
			writeFileSync(join(dir, 'run.sh'), harness);
			const run = spawnSync('bash', ['-e', '-o', 'pipefail', 'run.sh'], {
				cwd: dir,
				encoding: 'utf8',
				env: { ...process.env, PROJECT: 'vue', SHARD: '4', TOTAL: '10', SPECS: '' },
			});
			assert.equal(run.status, retryStatus, run.stderr || run.error?.message);
			assert.equal(
				readFileSync(join(dir, 'ci-evidence/attempt-1/test-results/trace.zip'), 'utf8'),
				'trace-1',
			);
			assert.equal(
				readFileSync(join(dir, 'ci-evidence/attempt-1/playwright-report/index.html'), 'utf8'),
				'report-1',
			);
			const commands = readFileSync(join(dir, 'commands.txt'), 'utf8').trim().split('\n');
			assert.equal(commands.length, 2);
			assert.ok(!commands[0].includes('--last-failed'));
			assert.ok(commands[1].includes('--last-failed'));
			if (retryStatus !== 0)
				assert.equal(
					readFileSync(join(dir, 'ci-evidence/attempt-2/test-results/trace.zip'), 'utf8'),
					'trace-2',
				);
		} finally {
			rmSync(dir, { recursive: true, force: true });
		}
	}
	assert.match(browserJob, /- name: Preserve the report\s+if: always\(\)/u);
	assert.match(browserJob, /viewers\/pptx\/ci-evidence\//u);
});
