import { fork } from 'node:child_process';
import { mkdir, mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

import { build } from 'tsup';

import type { SecurityCase } from './case';

const here = fileURLToPath(new URL('.', import.meta.url));
const coreRoot = join(here, '..', '..');
const HARNESS = new URL('./harness.mjs', import.meta.url);
/** Loading a case module is not part of its budget, but a load that never ends still fails. */
const LOAD_LIMIT_MS = 60_000;

/** What the harness reports for a case that finished within its limits. */
interface Done {
	value: unknown;
	error: string | undefined;
	ms: number;
	heapMb: number;
	network: string[];
}

/** What happened to one case in its process. */
export type Outcome =
	| ({ kind: 'done' } & Done)
	| { kind: 'over-time'; limitMs: number }
	| { kind: 'over-heap'; limitMb: number }
	| { kind: 'crashed'; error: string };

export interface CaseBundles {
	/** The bundle of each case module, by module name. */
	readonly files: ReadonlyMap<string, string>;
	readonly dispose: () => Promise<void>;
}

/**
 * Bundle the case modules for Node, keeping packages external. The output goes under the core's
 * `node_modules` so those packages resolve exactly as they do for the tests.
 */
export async function bundleCases(modules: ReadonlyMap<string, string>): Promise<CaseBundles> {
	const cache = join(coreRoot, 'node_modules', '.cache');
	await mkdir(cache, { recursive: true });
	const outDir = await mkdtemp(join(cache, 'security-corpus-'));
	await build({
		entry: Object.fromEntries(modules),
		outDir,
		format: 'esm',
		platform: 'node',
		target: 'node22',
		outExtension: () => ({ js: '.mjs' }),
		tsconfig: join(coreRoot, 'tsconfig.json'),
		config: false,
		dts: false,
		sourcemap: false,
		silent: true,
		skipNodeModulesBundle: true,
		noExternal: [/^ooxml-core\//],
	});
	const files = new Map([...modules.keys()].map((name) => [name, join(outDir, `${name}.mjs`)]));
	return { files, dispose: () => rm(outDir, { recursive: true, force: true }) };
}

/** Run case `index` of `bundle` in a child process held to the case's time and heap limits. */
export function runCase(bundle: string, index: number, limits: SecurityCase['limits']) {
	return new Promise<Outcome>((resolve) => {
		const child = fork(HARNESS, [pathToFileURL(bundle).href, String(index)], {
			execArgv: [`--max-old-space-size=${limits.heapMb}`],
			serialization: 'advanced',
			stdio: ['ignore', 'ignore', 'pipe', 'ipc'],
		});
		let stderr = '';
		child.stderr?.on('data', (chunk: Buffer) => {
			stderr = (stderr + chunk.toString()).slice(-4_000);
		});
		let settled = false;
		const settle = (outcome: Outcome) => {
			if (settled) return;
			settled = true;
			clearTimeout(timer);
			child.kill();
			resolve(outcome);
		};
		let timer = setTimeout(
			() => settle({ kind: 'crashed', error: 'the case module did not load' }),
			LOAD_LIMIT_MS,
		);
		child.on('message', (message: { type: 'ready' } | ({ type: 'done' } & Done)) => {
			if (message.type === 'ready') {
				clearTimeout(timer);
				timer = setTimeout(() => settle({ kind: 'over-time', limitMs: limits.ms }), limits.ms);
				return;
			}
			const { type: _type, ...done } = message;
			settle({ kind: 'done', ...done });
		});
		child.on('error', (error) => settle({ kind: 'crashed', error: error.stack ?? error.message }));
		// V8 aborts the process when the heap limit is reached; its last words name the cause.
		child.on('close', (code, signal) =>
			settle(
				/heap out of memory|Reached heap limit/u.test(stderr)
					? { kind: 'over-heap', limitMb: limits.heapMb }
					: { kind: 'crashed', error: `exited with ${signal ?? code}: ${stderr.trim()}` },
			),
		);
	});
}
