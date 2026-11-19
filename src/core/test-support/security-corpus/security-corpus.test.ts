import { fileURLToPath } from 'node:url';

import { afterAll, beforeAll, describe, it } from 'vitest';

import type { SecurityCase } from './case';
import { type CaseBundles, type Outcome, bundleCases, runCase } from './run-case';

declare global {
	interface ImportMeta {
		glob<T>(pattern: string, options: { eager: true }): Record<string, T>;
	}
}

/**
 * Every fixed security defect, as the smallest generated input that shows its class: one module
 * per area under `cases/`, each exporting `cases`. Each case runs in a process of its own with its
 * time and heap limits enforced and the network refused, so a regression fails here as a typed
 * result, a timeout or an exhausted heap instead of taking the test process with it. The glob is
 * a static import for Vite, so `vitest --changed` runs this file when code a case reaches changes.
 */
const MODULES = new Map(
	Object.entries(
		import.meta.glob<{ cases: readonly SecurityCase[] }>('./cases/*.ts', { eager: true }),
	).map(([path, module]) => [/([\w-]+)\.ts$/u.exec(path)![1]!, module.cases]),
);

/** Cases running at once: most of a case is starting its process, but CI's cores are shared. */
const CONCURRENT_CASES = 2;
/** Set to print each case's time and heap, for choosing its limits. */
const REPORT = Boolean(process.env['SECURITY_CORPUS_REPORT']);

let bundles: CaseBundles;
let running = 0;
const waiting: (() => void)[] = [];

async function inTurn(run: () => Promise<Outcome>): Promise<Outcome> {
	if (running >= CONCURRENT_CASES) await new Promise<void>((resolve) => waiting.push(resolve));
	running++;
	try {
		return await run();
	} finally {
		running--;
		waiting.shift()?.();
	}
}

beforeAll(async () => {
	const start = performance.now();
	bundles = await bundleCases(
		new Map(
			[...MODULES.keys()].map((name) => [
				name,
				fileURLToPath(new URL(`./cases/${name}.ts`, import.meta.url)),
			]),
		),
	);
	if (REPORT) console.log(`bundled in ${(performance.now() - start).toFixed(0)} ms`);
}, 60_000);

afterAll(() => bundles?.dispose());

for (const [module, cases] of MODULES) {
	describe.concurrent(module, () => {
		cases.forEach((entry, index) => {
			it(entry.name, { timeout: 120_000 }, async ({ expect }) => {
				const outcome = await inTurn(() =>
					runCase(bundles.files.get(module)!, index, entry.limits),
				);
				if (REPORT && outcome.kind === 'done')
					console.log(
						`${module} #${index}: ${outcome.ms.toFixed(0)} ms, ${outcome.heapMb.toFixed(0)} MB`,
					);
				const { kind } = outcome;
				const seen =
					kind === 'done'
						? { kind, error: outcome.error, network: outcome.network, value: outcome.value }
						: outcome;
				expect(seen).toEqual({
					kind: 'done',
					error: undefined,
					network: [],
					value: entry.expected,
				});
			});
		});
	});
}
