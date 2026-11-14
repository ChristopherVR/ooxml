import { readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';

/**
 * Safari and Firefox size a wrapping column flex container (`flex-direction: column` with
 * `flex-wrap: wrap`) as one column wide: the columns after the first spill over the next ribbon
 * group, or are clipped, while Chromium widens the box to fit them. The PowerPoint Home tab's
 * Drawing group drew its Shape Fill column over the Editing group in Safari that way. Ribbon rows
 * that stack commands into columns use a grid instead; this guards every stylesheet of the
 * package, `.css` files and the CSS held in TypeScript strings alike.
 */
const ROOT = path.join(import.meta.dirname, '..');
const COLUMN = /flex-direction\s*:\s*column(?!-)|flex-flow\s*:[^;{}]*\bcolumn\b(?!-)/u;
const WRAP = /flex-wrap\s*:\s*wrap(?!-)|flex-flow\s*:[^;{}]*\bwrap\b(?!-)/u;

function sources(): string[] {
	return readdirSync(ROOT, { recursive: true })
		.map((file) => String(file).replaceAll('\\', '/'))
		.filter((file) => /\.(css|ts)$/u.test(file) && !/\.(test|spec)\.ts$/u.test(file));
}

/** Declaration blocks of a stylesheet (the text between a `{` and the next brace). */
function columnWrapRules(text: string): string[] {
	return text
		.split('{')
		.slice(1)
		.map((block) => block.split('}')[0] ?? '')
		.filter((body) => COLUMN.test(body) && WRAP.test(body))
		.map((body) => body.trim().replaceAll(/\s+/gu, ' '));
}

describe('ribbon column stacks', () => {
	it('recognises a wrapping column flex container', () => {
		expect(columnWrapRules('.a { flex-direction: column; flex-wrap: wrap; }')).toHaveLength(1);
		expect(columnWrapRules('.a{flex-flow:column wrap}')).toHaveLength(1);
		expect(columnWrapRules('.a { flex-direction: column; flex-wrap: nowrap; }')).toEqual([]);
		expect(columnWrapRules('.a { flex-direction: row; flex-wrap: wrap; }')).toEqual([]);
	});

	it('never lays out a ribbon with a wrapping column flex container', () => {
		const offenders = sources().flatMap((file) =>
			columnWrapRules(readFileSync(path.join(ROOT, file), 'utf8')).map(
				(rule) => `${file}: ${rule}`,
			),
		);
		expect(offenders).toEqual([]);
	});
});
