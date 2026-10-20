import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { argumentBatches, FORMATTER_CHECK } from './verify-arguments.mjs';

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

test('formatter batches allow skipped formats while still rejecting unformatted source', () => {
	const dir = mkdtempSync(join(tmpdir(), 'ooxml-verify-format-'));
	try {
		const skipped = join(dir, 'control.svelte');
		const source = join(dir, 'control.ts');
		writeFileSync(skipped, '<button>Example</button>');
		writeFileSync(source, 'const value= {a:1,b:2};');
		const run = (files) =>
			spawnSync(process.platform === 'win32' ? 'bash.exe' : 'bash', ['-s'], {
				encoding: 'utf8',
				input: argumentBatches(FORMATTER_CHECK, files).join('\n') + '\n',
			});
		const ignored = run([skipped]);
		assert.equal(ignored.status, 0, ignored.stdout + ignored.stderr);
		const invalid = run([skipped, source]);
		assert.equal(invalid.status, 1, invalid.stdout + invalid.stderr);
	} finally {
		rmSync(dir, { recursive: true, force: true });
	}
});
