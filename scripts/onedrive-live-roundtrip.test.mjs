import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { parseRoundtripArgs, readEnvKey, resolveToken } from './onedrive-live-roundtrip.mjs';

const script = fileURLToPath(new URL('./onedrive-live-roundtrip.mjs', import.meta.url));

test('parses options and rejects unknown ones', () => {
	assert.deepEqual(parseRoundtripArgs(['--env', 'a.env', '--restore']), {
		env: 'a.env',
		linkEnv: 'OOXML_ONEDRIVE_SHARE_URL',
		restore: true,
	});
	assert.throws(() => parseRoundtripArgs(['--env']), /needs a value/u);
	assert.throws(() => parseRoundtripArgs(['--nope']), /Unknown option/u);
});

test('reads one key from .env text', () => {
	const text = '# c\nOTHER=1\nexport GRAPH_TOKEN="Bearer abc.def"\r\n';
	assert.equal(readEnvKey(text, 'GRAPH_TOKEN'), '"Bearer abc.def"');
	assert.equal(readEnvKey(text, 'MISSING'), undefined);
	assert.equal(resolveToken({}, text), 'abc.def');
	assert.equal(resolveToken({ GRAPH_TOKEN: ' xyz ' }, text), 'xyz');
	assert.equal(resolveToken({}, undefined), '');
});

test('skips with exit code 0 and no output of secrets when no token exists', () => {
	const env = { ...process.env };
	delete env.GRAPH_TOKEN;
	delete env.OOXML_ONEDRIVE_ACCESS_TOKEN;
	const run = spawnSync(process.execPath, [script, '--env', 'does-not-exist.env'], {
		env,
		encoding: 'utf8',
	});
	assert.equal(run.status, 0);
	assert.match(run.stdout, /SKIP/u);
});
