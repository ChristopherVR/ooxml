// Declarations (tsc) and the ESM bundle (tsup), built concurrently into dist.
import { spawn } from 'node:child_process';
import { rmSync } from 'node:fs';
import { rewriteEsmDeclarationImports } from '../../../scripts/esm-declarations.mjs';

const run = (cmd, args) =>
	new Promise((resolve, reject) => {
		const child = spawn(cmd, args, { stdio: 'inherit', shell: process.platform === 'win32' });
		child.on('exit', (code) =>
			code === 0 ? resolve() : reject(new Error(`${cmd} ${args.join(' ')} exited with ${code}`)),
		);
	});

rmSync('dist', { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
// Strict embedded surfaces consume the relaxed PowerPoint declarations, never its source.
rmSync('.types-pptx', { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
await run('tsc', ['-p', 'tsconfig.pptx.json', '--outDir', '.types-pptx']);
await Promise.all([
	(async () => {
		// Both projects emit shared declarations into `dist/`; run them sequentially to avoid races.
		await run('tsc', ['-p', 'tsconfig.pptx.json']);
		await run('tsc', ['-p', 'tsconfig.build.json']);
	})(),
	run('tsup', ['--config', 'tsup.pptx.config.ts']),
	run('tsup', ['--config', 'tsup.config.ts']),
]);
rewriteEsmDeclarationImports('dist');
