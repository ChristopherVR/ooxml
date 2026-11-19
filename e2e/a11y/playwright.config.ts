/**
 * The accessibility gates, run against the built suite (`bun run build:suite`) and the five
 * PowerPoint demos (`bun e2e-build.ts` in viewers/pptx with PPTX_E2E_PROJECTS listing them).
 *
 *   bun run test                                  # everything
 *   bun run test -- --project=suite               # the suite only, no PowerPoint demos
 *   A11Y_PORT=4500 bun run test                   # another port range, for a second worktree
 *
 * The `suite` project serves `suite-dist` on A11Y_PORT; each `pptx-<binding>` project serves that
 * demo's `dist` on the ports after it. A run that selects no PowerPoint project starts no demo.
 */
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { defineConfig, devices } from '@playwright/test';
import { browserPerformance } from '../../scripts/browser-performance.mjs';

const ROOT = fileURLToPath(new URL('../../', import.meta.url));
const PORT = Number(process.env.A11Y_PORT ?? 4410);
const isCI = Boolean(process.env.CI);
const performance = browserPerformance();

export const BINDINGS = ['react', 'vue', 'angular', 'vanilla', 'svelte'] as const;
const demoPort = (binding: (typeof BINDINGS)[number]) => PORT + 1 + BINDINGS.indexOf(binding);

/** The `--project` values of the run; workers inherit them through the environment. */
function selectedProjects(): string[] | undefined {
	const found: string[] = [];
	process.argv.forEach((arg, at) => {
		const inline = /^(?:--project|-p)=(.+)$/u.exec(arg)?.[1];
		const next = process.argv[at + 1];
		if (inline) found.push(inline);
		else if ((arg === '--project' || arg === '-p') && next) found.push(next);
	});
	if (found.length) process.env.A11Y_PROJECTS = found.join(',');
	return process.env.A11Y_PROJECTS?.split(',');
}
const selected = selectedProjects();
const runs = (project: string) => !selected || selected.includes(project);

function requireBuilt(path: string, how: string): void {
	if (!existsSync(`${ROOT}${path}`))
		throw new Error(`The accessibility gates need ${path}: ${how}.`);
}

const servers = [];
if (runs('suite')) {
	requireBuilt('suite-dist/index.html', 'run `bun run build:suite` at the repository root');
	servers.push({
		command: `node scripts/preview-site.mjs ${PORT}`,
		cwd: ROOT,
		url: `http://127.0.0.1:${PORT}/?suite`,
	});
}
for (const binding of BINDINGS) {
	if (!runs(`pptx-${binding}`)) continue;
	requireBuilt(
		`demos/pptx/demo-${binding}/dist/index.html`,
		'run `bun e2e-build.ts` in viewers/pptx with PPTX_E2E_PROJECTS=react,vue,angular,vanilla,svelte',
	);
	servers.push({
		command: `bun x vite preview --host 127.0.0.1 --port ${demoPort(binding)} --strictPort`,
		cwd: `${ROOT}demos/pptx/demo-${binding}`,
		url: `http://127.0.0.1:${demoPort(binding)}`,
	});
}

export default defineConfig({
	testDir: '.',
	// Only the specs: check-baseline.test.mjs is a node:test file (`bun run test:unit`).
	testMatch: '**/*.spec.ts',
	timeout: 90_000,
	expect: { timeout: 10_000 },
	fullyParallel: true,
	workers: performance.workers,
	forbidOnly: isCI,
	retries: isCI ? 1 : 0,
	reporter: isCI ? [['list'], ['github'], ['html', { open: 'never' }]] : [['list']],
	use: {
		...devices['Desktop Chrome'],
		trace: performance.trace,
		actionTimeout: 15_000,
		...(performance.hosted && {
			launchOptions: { args: ['--disable-dev-shm-usage', '--disable-gpu'] },
		}),
	},
	projects: [
		{
			name: 'suite',
			testIgnore: '**/pptx-bindings.spec.ts',
			use: { baseURL: `http://127.0.0.1:${PORT}` },
		},
		...BINDINGS.map((binding) => ({
			name: `pptx-${binding}`,
			testMatch: '**/pptx-bindings.spec.ts',
			use: { baseURL: `http://127.0.0.1:${demoPort(binding)}` },
		})),
	],
	webServer: servers.map((server) => ({
		...server,
		reuseExistingServer: !isCI,
		timeout: 120_000,
		stdout: 'ignore' as const,
		stderr: 'pipe' as const,
	})),
});
