import { execFileSync } from 'node:child_process';
import { existsSync, realpathSync } from 'node:fs';
import { delimiter, dirname, resolve } from 'node:path';

// npm.cmd is not executable by execFileSync on Windows. Run npm's JavaScript
// entry through Node, retaining argument boundaries without invoking a shell.
export function npmCliPath(env = process.env) {
	if (env.npm_execpath && existsSync(env.npm_execpath)) return env.npm_execpath;
	const nodeDirectory = dirname(process.execPath);
	const directories = [nodeDirectory, ...(env.PATH ?? env.Path ?? '').split(delimiter)];
	for (const directory of directories) {
		if (!directory) continue;
		for (const candidate of [
			resolve(directory, 'node_modules/npm/bin/npm-cli.js'),
			resolve(directory, '../lib/node_modules/npm/bin/npm-cli.js'),
		]) {
			if (existsSync(candidate)) return candidate;
		}
		const executable = resolve(directory, 'npm');
		if (process.platform !== 'win32' && existsSync(executable)) {
			const target = realpathSync(executable);
			if (target.endsWith('/npm-cli.js')) return target;
		}
	}
	throw new Error(
		'Cannot locate npm CLI. Run this script through npm, or install npm alongside Node.',
	);
}

export function runNpm(args, options = {}) {
	return execFileSync(process.execPath, [npmCliPath(options.env ?? process.env), ...args], {
		...options,
		shell: false,
	});
}
