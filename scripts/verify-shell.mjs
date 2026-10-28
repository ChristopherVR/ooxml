import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';

/** WSL's bash.exe drops Windows environment variables and runs a different Bun installation. */
export function verificationShell() {
	if (process.env.OOXML_BASH) return process.env.OOXML_BASH;
	if (process.platform !== 'win32') return 'bash';
	const candidates = [];
	try {
		for (const git of execFileSync('where.exe', ['git.exe'], { encoding: 'utf8' })
			.trim()
			.split(/\r?\n/))
			candidates.push(resolve(dirname(git), '../bin/bash.exe'));
	} catch {
		// Also support Git installations not currently on PATH.
	}
	for (const root of [process.env.ProgramFiles, process.env['ProgramFiles(x86)']])
		if (root) candidates.push(join(root, 'Git/bin/bash.exe'));
	if (process.env.LOCALAPPDATA)
		candidates.push(join(process.env.LOCALAPPDATA, 'Programs/Git/bin/bash.exe'));
	const shell = candidates.find(existsSync);
	if (!shell) throw new Error('Install Git for Windows or set OOXML_BASH to its bash.exe');
	return shell;
}
