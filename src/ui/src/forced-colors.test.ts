// @vitest-environment node
import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

/**
 * Forced-colours (Windows contrast themes) replace every author colour that is not a system
 * colour, so a selected or pressed state drawn only in accent colours disappears. These checks
 * read every stylesheet of the shared elements and the product editors.
 */
const ROOT = fileURLToPath(new URL('.', import.meta.url));
function stylesheets(directory: string): [string, string][] {
	return readdirSync(directory, { withFileTypes: true }).flatMap((entry): [string, string][] => {
		const path = join(directory, entry.name);
		if (entry.isDirectory()) return stylesheets(path);
		return entry.name.endsWith('.css') ? [[relative(ROOT, path), readFileSync(path, 'utf8')]] : [];
	});
}

/** Attribute and class spellings of a selected, current, checked or pressed state. */
const STATE =
	/aria-selected=['"]true['"]|aria-current(?:=['"][a-z]+['"])?|aria-pressed=['"]true['"]|aria-checked=['"]true['"]|\[selected\]|\[active\]|\[current\]|\[pressed\]|\.(?:active|selected|current)\b|:checked/g;
const PAINTS = /(?:^|[;\s])(?:background(?:-color)?|color|border(?:-[a-z]+)?-color|box-shadow)\s*:/;

interface Rule {
	selector: string;
	body: string;
	forced: boolean;
}

function rules(source: string): Rule[] {
	const css = source.replace(/\/\*[\s\S]*?\*\//g, '');
	const forced: [number, number][] = [];
	for (const match of css.matchAll(/@media[^{]*forced-colors:\s*active[^{]*\{/g)) {
		let depth = 1;
		let end = match.index + match[0].length;
		while (depth && end < css.length) {
			const char = css[end++];
			if (char === '{') depth++;
			else if (char === '}') depth--;
		}
		forced.push([match.index, end]);
	}
	return [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)].map((match) => ({
		selector: match[1]!.replace(/@media[^{]*$/, '').trim(),
		body: match[2]!,
		forced: forced.some(([start, end]) => match.index >= start && match.index < end),
	}));
}

describe('forced colours', () => {
	const files = stylesheets(ROOT);

	it('reads the stylesheets', () => {
		expect(files.length).toBeGreaterThan(100);
	});

	it('gives every coloured selected or pressed state a forced-colours rule', () => {
		const missing: string[] = [];
		for (const [file, css] of files) {
			const all = rules(css);
			const covered = new Set(
				all.filter((r) => r.forced).flatMap((r) => r.selector.match(STATE) ?? []),
			);
			const needed = new Set(
				all
					.filter((r) => !r.forced && PAINTS.test(r.body) && !/display\s*:\s*none/.test(r.body))
					.flatMap((r) => r.selector.match(STATE) ?? []),
			);
			for (const state of needed) if (!covered.has(state)) missing.push(`${file} ${state}`);
		}
		expect(missing).toEqual([]);
	});

	it('removes the text backplate wherever text is drawn in HighlightText', () => {
		// With the backplate, the browser paints Canvas behind HighlightText, which hides the label.
		const backplated = files.flatMap(([file, css]) =>
			rules(css)
				.filter(
					(r) =>
						r.forced &&
						/(?:^|[;\s])color\s*:\s*(?:HighlightText|SelectedItemText)/.test(r.body) &&
						!/forced-color-adjust\s*:\s*none/.test(r.body),
				)
				.map((r) => `${file} ${r.selector}`),
		);
		expect(backplated).toEqual([]);
	});
});
