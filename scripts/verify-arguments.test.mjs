import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { test } from 'node:test';
import { argumentBatches } from './verify-arguments.mjs';

test('large file lists reach bash unchanged across bounded commands', () => {
	const files = Array.from(
		{ length: 400 },
		(_, index) =>
			`src/path with spaces/${index}-quote'-$HOME-$(echo altered)-\`echo changed\`-file.ts`,
	);
	const commands = argumentBatches("printf '%s\\0'", files);
	assert.ok(commands.length > 1);
	const actual = commands.flatMap((command) => {
		assert.ok(command.length <= 4000);
		return execFileSync(process.platform === 'win32' ? 'bash.exe' : 'bash', ['-s'], {
			encoding: 'utf8',
			input: command + '\n',
		})
			.split('\0')
			.slice(0, -1);
	});
	assert.deepEqual(actual, files);
});

test('empty lists run nothing and oversized single arguments fail explicitly', () => {
	assert.deepEqual(argumentBatches('formatter', []), []);
	assert.throws(() => argumentBatches('formatter', ['x'.repeat(4000)]), /exceeds command limit/);
});
