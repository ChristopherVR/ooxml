// @vitest-environment node
import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { cssInTypeScript, specificity } from './forced-colors-audit';

/**
 * Forced-colours (Windows contrast themes) replace every author colour that is not a system
 * colour, so a selected or pressed state drawn only in accent colours disappears. These checks
 * read every stylesheet of the shared elements and the product editors, including the CSS that
 * modules keep in template strings.
 */
const ROOT = fileURLToPath(new URL('.', import.meta.url));
function stylesheets(directory: string): [string, string][] {
	return readdirSync(directory, { withFileTypes: true }).flatMap((entry): [string, string][] => {
		const path = join(directory, entry.name);
		if (entry.isDirectory()) return stylesheets(path);
		const name = relative(ROOT, path);
		if (entry.name.endsWith('.css')) return [[name, readFileSync(path, 'utf8')]];
		if (entry.name.endsWith('.ts') && !/\.(test|spec)\.ts$|\.d\.ts$/.test(entry.name)) {
			const css = cssInTypeScript(readFileSync(path, 'utf8'));
			return css ? [[name, css]] : [];
		}
		return [];
	});
}

/**
 * Attribute and class spellings of a selected, current, checked or pressed state: the ARIA states,
 * `:checked`, and a class that ends in the state's name (`.active`, `.is-active`, `.xst-active`).
 */
const STATE =
	/aria-selected=['"]true['"]|aria-current(?:=['"][a-z]+['"])?|aria-pressed=['"]true['"]|aria-checked=['"]true['"]|\[selected\]|\[active\]|\[current\]|\[pressed\]|\.[\w-]*(?<![\w-]in|\bin)(?:active|selected|current|pressed|checked)(?![\w-])|:checked/g;
const PAINTS = /(?:^|[;\s])(?:background(?:-color)?|color|border(?:-[a-z]+)?-color|box-shadow)\s*:/;

interface Rule {
	selectors: string[];
	body: string;
	forced: boolean;
}

/** Split a selector list on its top-level commas. */
function selectorList(list: string): string[] {
	const out: string[] = [];
	let depth = 0;
	let start = 0;
	for (let i = 0; i < list.length; i++) {
		const char = list[i];
		if (char === '(' || char === '[') depth++;
		else if (char === ')' || char === ']') depth--;
		else if (char === ',' && !depth) {
			out.push(list.slice(start, i).trim());
			start = i + 1;
		}
	}
	out.push(list.slice(start).trim());
	return out.filter(Boolean);
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
		selectors: selectorList(match[1]!.replace(/@media[^{]*$/, '').trim()),
		body: match[2]!,
		forced: forced.some(([start, end]) => match.index >= start && match.index < end),
	}));
}

/** Whether a forced rule's paint wins over any normal rule: `!important` on a colour property. */
const importantPaint = (body: string) =>
	/(?:background(?:-color)?|color|border(?:-[a-z]+)?-color|outline(?:-color)?)\s*:[^;]*!important/.test(
		body,
	);

const compare = (a: readonly number[], b: readonly number[]) =>
	a[0]! - b[0]! || a[1]! - b[1]! || a[2]! - b[2]!;

describe('forced colours', () => {
	const files = stylesheets(ROOT);

	it('reads the stylesheets and the CSS held in modules', () => {
		expect(files.filter(([file]) => file.endsWith('.css')).length).toBeGreaterThan(100);
		expect(files.filter(([file]) => file.endsWith('.ts')).length).toBeGreaterThan(20);
	});

	it('gives every coloured selected or pressed state a forced-colours rule that wins', () => {
		const missing: string[] = [];
		for (const [file, css] of files) {
			const all = rules(css);
			// The strongest forced-colours selector for each state spelling.
			const covered = new Map<string, number[]>();
			for (const rule of all.filter((r) => r.forced)) {
				for (const selector of rule.selectors) {
					const weight = importantPaint(rule.body) ? [Infinity, 0, 0] : specificity(selector);
					for (const state of selector.match(STATE) ?? []) {
						const best = covered.get(state);
						if (!best || compare(weight, best) > 0) covered.set(state, weight);
					}
				}
			}
			for (const rule of all) {
				if (rule.forced || !PAINTS.test(rule.body) || /display\s*:\s*none/.test(rule.body))
					continue;
				for (const selector of rule.selectors) {
					// A state inside `:not(...)` styles the unselected element, which needs no rule.
					const positive = selector.replace(/:not\((?:[^()]|\([^()]*\))*\)/g, '');
					for (const state of new Set(positive.match(STATE) ?? [])) {
						const best = covered.get(state);
						if (!best) missing.push(`${file} ${state}: no forced-colours rule`);
						else if (compare(best, specificity(selector)) < 0)
							missing.push(`${file} ${selector}: outranks its forced-colours rule`);
					}
				}
			}
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
				.map((r) => `${file} ${r.selectors.join(', ')}`),
		);
		expect(backplated).toEqual([]);
	});
});

describe('the forced-colours audit', () => {
	it('weighs selectors by CSS specificity', () => {
		expect(specificity('.tab[aria-selected="true"]')).toEqual([0, 2, 0]);
		expect(specificity(":host([plain]) .tab[aria-selected='true']")).toEqual([0, 4, 0]);
		expect(specificity(':where(.a, #b) li.is-active')).toEqual([0, 1, 1]);
		expect(specificity(':is(.a, #b) > button:not(.x)::before')).toEqual([1, 1, 2]);
		expect(specificity('::slotted([slot="after"])')).toEqual([0, 1, 1]);
	});

	it('reads CSS out of template strings and skips the code around them', () => {
		const source =
			"const a = `.x{color:${tok('--y')}}`;\nconst b = 'not css {}';\n// `.c{color:red}`\nconst c = html`<p class=${`z`}></p>`;";
		const css = cssInTypeScript(source);
		expect(css).toContain('.x{color:x}');
		expect(css).not.toContain('not css');
		expect(css).not.toContain('.c{');
	});
});
