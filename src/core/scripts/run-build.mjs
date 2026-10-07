import { spawn } from 'node:child_process';

/** Shared build command runner for full and format-specific builds. */
export const runBuild = (cmd, args) =>
	new Promise((resolve, reject) => {
		const child = spawn(cmd, args, { stdio: 'inherit', shell: process.platform === 'win32' });
		child.on('error', reject);
		child.on('exit', (code) =>
			code === 0 ? resolve() : reject(new Error(`${cmd} ${args.join(' ')} exited with ${code}`)),
		);
	});
