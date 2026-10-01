// Declarations (tsc) and the ESM bundle (tsup), built concurrently into dist.
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
await Promise.all([
	run('tsc', ['-p', 'tsconfig.build.json']),
	run('tsup', ['--config', 'tsup.config.ts']),
]);
