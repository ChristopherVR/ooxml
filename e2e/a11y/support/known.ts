/**
 * Gaps the products have today, each owned by the 2.0 work package that closes it. A check listed
 * in `expected-failures.json` runs as an expected failure: it keeps running, and it fails the
 * moment the gap closes, so the entry has to be deleted in the commit that fixes it. The list can
 * only shrink (`check-baseline.mjs` compares it with the base branch).
 */
import { test } from '@playwright/test';
import registry from '../expected-failures.json' with { type: 'json' };

export interface KnownGap {
	/** The work package that closes the gap. */
	owner: string;
	/** The audit finding it comes from. */
	finding: string;
	why: string;
}

const gaps: Record<string, KnownGap> = registry.gaps;

/** Marks the running test as an expected failure when `key` is a known gap. */
export function knownGap(key: string): void {
	const gap = gaps[key];
	test.info().annotations.push({ type: 'a11y-check', description: key });
	if (gap) test.fail(true, `${gap.owner} (${gap.finding}): ${gap.why}`);
}

/** The known gap for `key`, for checks that collect several results before asserting. */
export const gapFor = (key: string): KnownGap | undefined => gaps[key];
