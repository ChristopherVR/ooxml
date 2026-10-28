import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { test } from 'node:test';
import { verificationShell } from './verify-shell.mjs';

test('verification inherits the local worker budget instead of crossing into WSL', () => {
	const result = execFileSync(verificationShell(), ['-c', 'printf "%s" "$OOXML_E2E_WORKERS"'], {
		env: { ...process.env, OOXML_E2E_WORKERS: '8' },
		encoding: 'utf8',
	});
	assert.equal(result, '8');
});
