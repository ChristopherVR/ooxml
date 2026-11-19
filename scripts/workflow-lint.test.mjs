import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';

import { parse } from 'yaml';

import {
	lintRepository,
	lintWorkflow,
	lockedCommands,
	runnerInvocations,
	writeScopes,
} from './workflow-lint.mjs';

const LOCK = `{
  "lockfileVersion": 1,
  "workspaces": {
    "": { "name": "root", "devDependencies": { "git-cliff": "2.14.2", "@playwright/test": "1.0.0", }, },
    "src/core": { "name": "core", "devDependencies": { "vitest": "4.0.0", }, },
  },
  "packages": {
    "git-cliff": ["git-cliff@2.14.2", "", { "bin": "lib/cli/cli.js" }, "sha512-x"],
    "@playwright/test": ["@playwright/test@1.0.0", "", { "bin": { "playwright": "cli.js" } }, "sha512-y"],
    "vitest": ["vitest@4.0.0", "", { "bin": { "vitest": "vitest.mjs" } }, "sha512-z"],
  },
}`;

/** A throwaway repository holding `files`, linted as a whole. */
function repository(files) {
	const root = mkdtempSync(join(tmpdir(), 'workflow-lint-'));
	for (const [path, text] of Object.entries({ 'bun.lock': LOCK, ...files })) {
		mkdirSync(join(root, path, '..'), { recursive: true });
		writeFileSync(join(root, path), text);
	}
	mkdirSync(join(root, '.github/actions'), { recursive: true });
	try {
		return lintRepository(root);
	} finally {
		rmSync(root, { recursive: true, force: true });
	}
}

const job = (permissions, steps) => `
permissions:
  contents: read
jobs:
  work:
    runs-on: ubuntu-latest
${permissions ? `    permissions:\n${permissions.replace(/^/gmu, '      ')}\n` : ''}    steps:
${steps.replace(/^/gmu, '      ')}
`;

test('the repository follows the policy', () => {
	assert.deepEqual(lintRepository(), []);
});

test('locked commands are the workspaces dependencies and their bins', () => {
	const locked = lockedCommands(LOCK);
	for (const name of ['git-cliff', '@playwright/test', 'playwright', 'vitest']) {
		assert.ok(locked.has(name), name);
	}
	assert.ok(!locked.has('pkg-pr-new'));
});

test('runner invocations name the package each runner fetches', () => {
	assert.deepEqual(
		runnerInvocations(
			'cd src/core && bunx vitest run\nbun x --bun playwright install\nnpx -y "tool@1"',
		),
		[
			{ runner: 'bunx', spec: 'vitest' },
			{ runner: 'bun x', spec: 'playwright' },
			{ runner: 'npx', spec: 'tool@1' },
		],
	);
	assert.deepEqual(runnerInvocations('npm exec --package=cowsay -- cowsay hi'), [
		{ runner: 'npm exec', spec: 'cowsay' },
	]);
});

test('bunx and npx of an unlocked package fail, a locked one passes', () => {
	const findings = repository({
		'.github/workflows/a.yml': job(
			undefined,
			'- run: bunx git-cliff --version\n- run: bunx pkg-pr-new publish\n- run: npx git-cliff@2.14.2\n- run: bunx "$TOOL"',
		),
	});
	assert.equal(findings.length, 3, findings.join('\n'));
	assert.match(findings[0], /pkg-pr-new.*not a dependency/u);
	assert.match(findings[1], /git-cliff@2\.14\.2.*names a version/u);
	assert.match(findings[2], /computed/u);
});

test('a job that can write runs no third-party action, tool or package script', () => {
	const findings = repository({
		'scripts/clean.mjs': "import { readFileSync } from 'node:fs';\nimport './helper.mjs';\n",
		'scripts/helper.mjs': 'export const x = 1;\n',
		'scripts/dirty.mjs': "import { x } from './helper.mjs';\nimport { parse } from 'yaml';\n",
		'scripts/spawns.mjs':
			"import { execFileSync } from 'node:child_process';\nexecFileSync('bunx', ['x']);\n",
		'.github/workflows/w.yml': job(
			'contents: write',
			[
				'- uses: actions/checkout@v7',
				'- run: node scripts/clean.mjs && git push origin HEAD:main',
				'- uses: oven-sh/setup-bun@abc',
				'- run: bun install',
				'- run: node scripts/dirty.mjs',
				'- run: node scripts/spawns.mjs',
				'- run: node_modules/.bin/git-cliff',
			].join('\n'),
		),
	});
	assert.equal(findings.length, 5, findings.join('\n'));
	assert.match(findings[0], /uses oven-sh\/setup-bun in a job whose token writes contents/u);
	assert.match(findings[1], /runs `bun`/u);
	assert.match(findings[2], /scripts\/dirty\.mjs imports `yaml`/u);
	assert.match(findings[3], /scripts\/spawns\.mjs runs `bunx`/u);
	assert.match(findings[4], /node_modules/u);
});

test('a local composite action is checked step by step', () => {
	const findings = repository({
		'.github/actions/setup/action.yml':
			'runs:\n  using: composite\n  steps:\n    - uses: actions/setup-node@v7\n    - run: bun install --frozen-lockfile\n      shell: bash\n',
		'.github/workflows/w.yml': job('issues: write', '- uses: ./.github/actions/setup'),
	});
	assert.equal(findings.length, 1, findings.join('\n'));
	assert.match(findings[0], /\.\/\.github\/actions\/setup runs `bun`.*writes issues/u);
});

test('read-only and id-token jobs may run third-party code, but get no other secret', () => {
	const findings = repository({
		'.github/workflows/w.yml': job(
			'contents: read\nid-token: write',
			[
				'- uses: sigstore/cosign-installer@abc',
				'- run: cosign sign-blob x\n  env:\n    GH_TOKEN: ${{ secrets.GITHUB_TOKEN }}',
				'- run: cosign sign-blob y\n  env:\n    HOOK: ${{ secrets.DISCORD_WEBHOOK }}',
			].join('\n'),
		),
	});
	assert.equal(findings.length, 1, findings.join('\n'));
	assert.match(findings[0], /receives DISCORD_WEBHOOK/u);
});

test('write scopes come from the job, else the workflow; a missing block may write', () => {
	assert.deepEqual(writeScopes({ contents: 'read', 'id-token': 'write' }), []);
	assert.deepEqual(writeScopes({ contents: 'write', issues: 'write' }), ['contents', 'issues']);
	assert.deepEqual(writeScopes('read-all'), []);
	assert.deepEqual(writeScopes('write-all'), ['write-all']);
	assert.equal(writeScopes(undefined).length, 1);
	const workflow = parse(
		'permissions: write-all\njobs:\n  a:\n    steps:\n      - uses: foo/bar@v1\n',
	);
	const findings = lintWorkflow(workflow, { file: 'w.yml', root: '.', locked: new Set() });
	assert.match(findings[0], /uses foo\/bar in a job whose token writes write-all/u);
});
