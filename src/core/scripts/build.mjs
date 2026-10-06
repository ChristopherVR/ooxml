// Builds every area. The tsc declarations for the strict areas, the tsup bundles (strict areas and
// pptx) and the pptx declaration bundle are independent, so they run concurrently; only the
// declaration merge needs all of them.
import { spawn } from 'node:child_process';
import { rmSync } from 'node:fs';

const run = (cmd, args) =>
	new Promise((resolve, reject) => {
		const child = spawn(cmd, args, { stdio: 'inherit', shell: process.platform === 'win32' });
		child.on('exit', (code) =>
			code === 0 ? resolve() : reject(new Error(`${cmd} ${args.join(' ')} exited with ${code}`)),
		);
	});

rmSync('dist', { recursive: true, force: true });
rmSync('.types-pptx', { recursive: true, force: true });
await Promise.all([
	run('tsc', ['-p', 'tsconfig.build.json']),
	run('tsup', ['--config', 'tsup.config.ts']),
	run('tsup', ['--config', 'tsup.pptx.config.ts']),
	run('tsdown', ['--config', 'tsdown.pptx.config.ts']),
]);
await run('node', ['scripts/pptx/merge-declarations.mjs']);
