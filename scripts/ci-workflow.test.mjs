import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';

const workflow = readFileSync(new URL('../.github/workflows/ci.yml', import.meta.url), 'utf8');
const browserJob = workflow.match(/^  pptx-e2e:\r?\n([\s\S]*?)(?=^  \S)/m)?.[1];

test('push plans start at the last green commit, so a newer push may cancel an older one', () => {
	const concurrency = workflow.match(/^concurrency:\r?\n([\s\S]*?)(?=^\S)/m)?.[1];
	assert.ok(concurrency);
	assert.match(concurrency, /github\.event_name == 'push' && 'push-main'/u);
	assert.match(
		concurrency,
		/cancel-in-progress: \$\{\{ github.event_name == 'pull_request' \|\| github.event_name == 'push' \}\}/u,
	);
	const planJob = workflow.match(/^  plan:\r?\n([\s\S]*?)(?=^  \S)/m)?.[1];
	assert.ok(planJob);
	assert.match(planJob, /actions: read/u);
	assert.match(planJob, /--green-base true/u);
	assert.doesNotMatch(workflow.replace(planJob, ''), /actions: read/u);
});

/** Runs the browser-suite step with a fake `bun` whose attempts behave as `outcomes` says. */
function runBrowserStep(outcomes) {
	const command = browserJob.match(
		/- name: Run the browser suite[\s\S]*?        run: \|\r?\n([\s\S]*?)(?=^      -)/m,
	)?.[1];
	assert.ok(command, 'the browser job has a runner command');
	const script = command.replace(/^          /gm, '');
	const dir = mkdtempSync(join(tmpdir(), 'ooxml-ci-retry-'));
	try {
		// Mimic Playwright clearing its output on each run: `crash` leaves no report, `fail` records
		// failed tests, `pass` succeeds. Shell functions keep this portable to Windows and Linux.
		const harness = `
outcomes="${outcomes.join(' ')}"
attempts=0
bun() {
  attempts=$((attempts + 1))
  printf '%s\\n' "$*" >> commands.txt
  rm -rf test-results playwright-report
  mkdir -p test-results playwright-report
  printf 'trace-%s' "$attempts" > test-results/trace.zip
  outcome=$(echo $outcomes | cut -d' ' -f$attempts)
  case "$outcome" in
    pass) return 0 ;;
    fail) printf '{"status":"failed","failedTests":["t"]}' > test-results/.last-run.json; return 1 ;;
    *) return 1 ;;
  esac
}
sleep() { :; }
${script}`;
		writeFileSync(join(dir, 'run.sh'), harness);
		const run = spawnSync('bash', ['-e', '-o', 'pipefail', 'run.sh'], {
			cwd: dir,
			encoding: 'utf8',
			env: { ...process.env, PROJECT: 'vue', SHARD: '4', TOTAL: '7', SPECS: '' },
		});
		const commands = readFileSync(join(dir, 'commands.txt'), 'utf8').trim().split('\n');
		const firstTrace = readFileSync(
			join(dir, 'ci-evidence/attempt-1/test-results/trace.zip'),
			'utf8',
		);
		return { status: run.status, commands, firstTrace, stderr: run.stderr };
	} finally {
		rmSync(dir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
	}
}

test('recorded browser failures run again alone; a crash runs the shard again', () => {
	assert.ok(browserJob);
	const flaky = runBrowserStep(['fail', 'pass']);
	assert.equal(flaky.status, 0, flaky.stderr);
	assert.equal(flaky.commands.length, 2);
	assert.doesNotMatch(flaky.commands[0], /--last-failed/u);
	assert.match(flaky.commands[1], /--last-failed/u);
	assert.equal(flaky.firstTrace, 'trace-1');

	const failed = runBrowserStep(['fail', 'fail']);
	assert.notEqual(failed.status, 0);
	assert.equal(failed.commands.length, 2);
	assert.match(failed.commands[1], /--last-failed/u);

	const crashed = runBrowserStep(['crash', 'pass']);
	assert.equal(crashed.status, 0, crashed.stderr);
	assert.equal(crashed.commands.length, 2);
	assert.doesNotMatch(crashed.commands[1], /--last-failed/u);
	assert.equal(crashed.firstTrace, 'trace-1');

	const twice = runBrowserStep(['crash', 'crash']);
	assert.notEqual(twice.status, 0);
	assert.equal(twice.commands.length, 2);
	assert.ok(twice.commands.every((command) => command.includes('--shard=4/7')));
	assert.match(browserJob, /- name: Preserve the report\s+if: always\(\)/u);
	assert.match(browserJob, /viewers\/pptx\/ci-evidence\//u);
});
