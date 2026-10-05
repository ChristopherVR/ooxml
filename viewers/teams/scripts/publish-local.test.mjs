import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';

import {
	buildPublishPlan,
	formatPlan,
	isConfirmed,
	isOtpError,
	loginState,
	parseArgs,
	tagCommands,
} from './publish-local.mjs';

test('parses the supported flags and repeatable --package', () => {
	assert.deepEqual(parseArgs([]), {
		yes: false,
		dryRun: false,
		allowDirty: false,
		skipBuild: false,
		help: false,
		packages: [],
	});
	const options = parseArgs([
		'--dry-run',
		'--yes',
		'--allow-dirty',
		'--skip-build',
		'--package',
		'server',
		'--package',
		'react',
	]);
	assert.equal(options.dryRun, true);
	assert.equal(options.yes, true);
	assert.equal(options.allowDirty, true);
	assert.equal(options.skipBuild, true);
	assert.deepEqual(options.packages, ['server', 'react']);
});

test('rejects unknown flags and unknown packages', () => {
	assert.throws(() => parseArgs(['--force']), /Unknown argument/);
	assert.throws(() => parseArgs(['--package', 'teams-viewer']), /--package needs one of/);
	assert.throws(() => parseArgs(['--package']), /--package needs one of/);
});

test('never accepts a one-time password in any spelling', () => {
	for (const arg of ['--otp', '--otp=123456', '-otp', '--OTP', '--npm-otp=1']) {
		assert.throws(() => parseArgs([arg]), /never accepts a one-time password/, arg);
	}
	const source = readFileSync(new URL('./publish-local.mjs', import.meta.url), 'utf8');
	assert.doesNotMatch(source, /['"]--otp['"]/u, 'the script must never pass --otp to npm');
});

test('detects npm asking for a second factor', () => {
	assert.equal(
		isOtpError(
			'npm error code EOTP\nnpm error This operation requires a one-time password from your authenticator.',
		),
		true,
	);
	assert.equal(isOtpError('npm ERR! code EOTP'), true);
	assert.equal(
		isOtpError(
			'You must provide a one-time pass. Upgrade your client to npm@latest in order to use 2FA.',
		),
		true,
	);
	assert.equal(isOtpError('This command requires two-factor authentication'), true);
	assert.equal(
		isOtpError('npm error 403 Forbidden - You do not have permission to publish "react"'),
		false,
	);
	assert.equal(isOtpError('npm error code E404'), false);
	assert.equal(isOtpError('+ openteams-server@0.1.0'), false);
});

test('reads npm whoami results', () => {
	assert.equal(loginState({ status: 0, stdout: 'christophervr\n', stderr: '' }), 'ok');
	assert.equal(
		loginState({
			status: 1,
			stdout: '',
			stderr: 'npm error code E401\nnpm error 401 Unauthorized',
		}),
		'login',
	);
	assert.equal(loginState({ status: 1, stdout: '', stderr: 'npm error code ENEEDAUTH' }), 'login');
	assert.equal(
		loginState({
			status: 1,
			stdout: '',
			stderr: 'npm error code ENOTFOUND\nnpm error network request failed',
		}),
		'error',
	);
});

test('only the word yes confirms', () => {
	assert.equal(isConfirmed('yes'), true);
	assert.equal(isConfirmed('  YES \n'), true);
	for (const answer of ['y', 'Y', 'yes please', '', undefined, 'no'])
		assert.equal(isConfirmed(answer), false);
});

test('the publish plan keeps plan order, filters by key and marks versions already on npm', () => {
	const plan = {
		order: ['react', 'vue', 'server'],
		packages: {
			react: {
				release: true,
				npm: 'openteams-react-viewer',
				dir: 'packages/react',
				version: '0.1.0',
			},
			vue: { release: false, npm: 'openteams-vue-viewer', dir: 'packages/vue', version: '0.1.0' },
			server: { release: true, npm: 'openteams-server', dir: 'server', version: '0.1.0' },
		},
	};
	const state = (name) => (name === 'openteams-server' ? 'exists' : 'missing');
	const rows = buildPublishPlan(plan, state);
	assert.deepEqual(
		rows.map((r) => [r.key, r.npm, r.version, r.onNpm]),
		[
			['react', 'openteams-react-viewer', '0.1.0', false],
			['server', 'openteams-server', '0.1.0', true],
		],
	);
	assert.deepEqual(
		buildPublishPlan(plan, state, ['server']).map((r) => r.key),
		['server'],
	);
	const table = formatPlan(rows);
	assert.match(table, /openteams-react-viewer@0\.1\.0\s+will be published/);
	assert.match(table, /openteams-server@0\.1\.0\s+already on npm/);
	assert.match(formatPlan([]), /Nothing to publish/);
	assert.deepEqual(tagCommands(rows.slice(0, 1)), [
		'git tag openteams-react-viewer@0.1.0',
		'git push origin openteams-react-viewer@0.1.0',
	]);
	assert.deepEqual(tagCommands([]), []);
});
