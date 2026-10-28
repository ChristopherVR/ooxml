export function browserPerformance(
	env?: NodeJS.ProcessEnv,
	platform?: string,
	cpus?: number,
): {
	hosted: boolean;
	workers: number;
	trace: 'off' | 'on-first-retry';
};
export function pptxBrowserCommand(e2e: { specs: string; projects: string[] }): string;
