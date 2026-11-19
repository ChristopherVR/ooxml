/**
 * Gaps the products have today, each owned by the 2.0 work package that closes it. A check listed
 * in `expected-failures.json` must still fail, and fail the way its entry says (`fails`): then it
 * passes, annotated as the known gap. It fails the moment the gap closes, so the entry has to be
 * deleted in the commit that fixes it, and it fails when it breaks some other way, so a gap never
 * hides an unrelated regression. The list can only shrink (`check-baseline.mjs` compares it with
 * the base branch).
 */
import { test } from '@playwright/test';
import registry from '../expected-failures.json' with { type: 'json' };

export interface KnownGap {
	/** The work package that closes the gap. */
	owner: string;
	/** The audit finding it comes from. */
	finding: string;
	why: string;
	/** Text the check's failure message contains when it fails because of this gap. */
	fails: string;
}

const gaps: Record<string, KnownGap> = registry.gaps;

/** Assertion messages carry terminal colours; the expected text is matched without them. */
const plain = (text: string) => text.replace(/\u001b\[[\d;]*m/gu, '');

/** Runs the check `key`, holding a known gap to the one failure its entry expects. */
export async function gate(key: string, check: () => unknown): Promise<void> {
	const info = test.info();
	info.annotations.push({ type: 'a11y-check', description: key });
	const gap = gaps[key];
	if (!gap) {
		await check();
		return;
	}
	try {
		await check();
	} catch (error) {
		const message = plain(error instanceof Error ? error.message : String(error));
		if (gap.fails && message.includes(gap.fails)) {
			info.annotations.push({
				type: 'known-gap',
				description: `${gap.owner} (${gap.finding}): ${gap.why}`,
			});
			return;
		}
		throw new Error(
			`${key} is a known gap that fails with "${gap.fails}", but it failed another way:\n${message}`,
			{ cause: error },
		);
	}
	throw new Error(
		`${key} passes now: ${gap.owner} closed ${gap.finding}. Delete its entry from expected-failures.json.`,
	);
}
