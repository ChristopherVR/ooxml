import { availableParallelism } from 'node:os';

/** Verification is strict locally too, but Linux runner workarounds are platform-specific. */
export function browserPerformance(
	env = process.env,
	platform = process.platform,
	cpus = availableParallelism(),
) {
	const hosted = Boolean(env.CI) && platform === 'linux' && env.OOXML_LOCAL_VERIFY !== '1';
	const requested = env.OOXML_E2E_WORKERS;
	if (requested && (!/^\d+$/.test(requested) || Number(requested) < 1))
		throw new Error('OOXML_E2E_WORKERS must be a positive integer');
	return {
		hosted,
		workers: requested
			? Number(requested)
			: hosted
				? 2
				: Math.max(1, Math.min(8, Math.floor(cpus / 2))),
		trace: env.CI ? 'on-first-retry' : 'off',
	};
}

/** Preserve the planner's spec selection and both collaboration host variants. */
export function pptxBrowserCommand(e2e) {
	const quote = (value) => `'${value.replaceAll("'", "'\\''")}'`;
	const specs = e2e.specs ? e2e.specs.split(' ').filter(Boolean).map(quote) : [];
	const projects = e2e.projects.flatMap((name) => [name, `${name}-headless`]);
	return ['bun run e2e --', ...specs, ...projects.map((name) => `--project=${quote(name)}`)].join(
		' ',
	);
}
