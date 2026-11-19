/**
 * axe-core in the page, over open shadow roots, with the WCAG 2.0 to 2.2 A and AA rules. Only
 * critical and serious violations gate; the baseline (`baseline/axe.json`) says how many nodes
 * each surface may still have per rule, and it can only shrink.
 */
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import type { Page } from '@playwright/test';
import baseline from '../baseline/axe.json' with { type: 'json' };

const require = createRequire(import.meta.url);
const SOURCE = readFileSync(require.resolve('axe-core/axe.min.js'), 'utf8');
const TAGS = ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa'];
export const GATED = new Set(['critical', 'serious']);

export interface Violation {
	rule: string;
	impact: string;
	nodes: number;
	help: string;
	targets: string[];
}

/** Runs axe on the whole page and returns its critical and serious violations. */
export async function axeViolations(
	page: Page,
	disable: readonly string[] = [],
): Promise<Violation[]> {
	await page.evaluate(SOURCE);
	const found = await page.evaluate(
		async ({ tags, disable }) => {
			const axe = (globalThis as unknown as { axe: { run: Function } }).axe;
			const result = await axe.run(document, {
				runOnly: { type: 'tag', values: tags },
				rules: Object.fromEntries(disable.map((id) => [id, { enabled: false }])),
				resultTypes: ['violations'],
			});
			return (
				result.violations as {
					id: string;
					impact: string | null;
					helpUrl: string;
					nodes: { target: unknown[] }[];
				}[]
			).map((violation) => ({
				rule: violation.id,
				impact: violation.impact ?? 'minor',
				nodes: violation.nodes.length,
				help: violation.helpUrl,
				targets: violation.nodes
					.slice(0, 5)
					.map((node) => node.target.flat().join(' >> ').slice(0, 160)),
			}));
		},
		{ tags: TAGS, disable: [...disable] },
	);
	return found.filter((violation) => GATED.has(violation.impact));
}

const allowed: Record<string, Record<string, number>> = baseline.allowed;

/** The violations above what the baseline allows for `key` (`<surface>/<mode>`). */
export function overBaseline(key: string, violations: readonly Violation[]): Violation[] {
	return violations.filter((violation) => violation.nodes > (allowed[key]?.[violation.rule] ?? 0));
}

/** Baseline entries for `key` that this run no longer needs, to report so they get lowered. */
export function slack(key: string, violations: readonly Violation[]): string[] {
	return Object.entries(allowed[key] ?? {})
		.filter(
			([rule, count]) =>
				(violations.find((violation) => violation.rule === rule)?.nodes ?? 0) < count,
		)
		.map(([rule, count]) => `${rule} (${count} allowed)`);
}
