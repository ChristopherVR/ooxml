#!/usr/bin/env node
// Hash a password for the server's users table (TEAMS_USERS_FILE or the `users` option).
//   openteams-hash-password                       prompts (hidden) when run in a terminal
//   echo -n 'secret' | openteams-hash-password    reads standard input otherwise
// The password is never taken from the command line (shell history, process list) and the output
// is only the hash. Paste it as the user's value: { "alice": "scrypt$16384$8$1$..." }.
import { hashPassword } from './passwords.mjs';

function readHidden(prompt) {
	process.stderr.write(prompt);
	const stdin = process.stdin;
	stdin.setRawMode(true);
	stdin.resume();
	stdin.setEncoding('utf8');
	return new Promise((resolve) => {
		let value = '';
		const onData = (chunk) => {
			for (const ch of chunk) {
				if (ch === '\r' || ch === '\n' || ch === '\u0004') {
					stdin.setRawMode(false);
					stdin.pause();
					stdin.off('data', onData);
					process.stderr.write('\n');
					return resolve(value);
				}
				if (ch === '\u0003') {
					stdin.setRawMode(false);
					process.stderr.write('\n');
					process.exit(130);
				}
				if (ch === '\u007f' || ch === '\b') value = value.slice(0, -1);
				else value += ch;
			}
		};
		stdin.on('data', onData);
	});
}

async function readStdin() {
	const chunks = [];
	for await (const chunk of process.stdin) chunks.push(chunk);
	return Buffer.concat(chunks)
		.toString('utf8')
		.replace(/\r?\n$/u, '');
}

async function runHashPassword() {
	try {
		let password;
		if (process.stdin.isTTY) {
			password = await readHidden('Password: ');
			if (password !== (await readHidden('Repeat: '))) throw new Error('passwords do not match');
		} else {
			password = await readStdin();
		}
		console.log(await hashPassword(password));
	} catch (error) {
		console.error(`openteams-hash-password: ${error.message}`);
		process.exit(1);
	}
}

// An npm bin is started through a link, so this file is always the entry point.
await runHashPassword();
