import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

// The generator is a build script, not part of the package; the test imports it directly.
// @ts-expect-error - plain JavaScript module without declarations.
import * as preflight from '../../../scripts/pptx-viewer-preflight.mjs';

interface PreflightModule {
	VIEWER_SCOPES: string[];
	OUTPUT: string;
	generate(): string;
	readInstalledPreflight(): { source: string; version: string; license: string };
	scopeSelector(selector: string): string[];
}

const { VIEWER_SCOPES, OUTPUT, generate, readInstalledPreflight, scopeSelector } =
	preflight as PreflightModule;
const SCOPE = `:where(${VIEWER_SCOPES.join(',')})`;
const committed = readFileSync(OUTPUT, 'utf8');

/** Every selector of every qualified rule, comments and declarations removed. */
function selectorsOf(css: string): string[] {
	const out: string[] = [];
	const stripped = css.replace(/\/\*[\s\S]*?\*\//gu, '');
	for (const match of stripped.matchAll(/([^{}]+)\{([^{}]*)\}/gu)) {
		const prelude = (match[1] ?? '').trim();
		if (!prelude || prelude.startsWith('@')) continue;
		out.push(...splitSelectors(prelude));
	}
	return out;
}

function splitSelectors(prelude: string): string[] {
	const parts: string[] = [];
	let depth = 0;
	let start = 0;
	for (let index = 0; index < prelude.length; index++) {
		const char = prelude[index];
		if (char === '(' || char === '[') depth++;
		else if (char === ')' || char === ']') depth--;
		else if (char === ',' && depth === 0) {
			parts.push(prelude.slice(start, index).trim());
			start = index + 1;
		}
	}
	parts.push(prelude.slice(start).trim());
	return parts.filter(Boolean);
}

describe('viewer-preflight.css', () => {
	it('is the scoped transform of the installed tailwindcss preflight', () => {
		// Regenerate with `node scripts/pptx-viewer-preflight.mjs` in src/ui after a Tailwind upgrade.
		expect(committed).toBe(generate());
	});

	it('keeps the Tailwind licence notice and names the source version', () => {
		const { version } = readInstalledPreflight();
		expect(committed).toContain(`tailwindcss@${version}`);
		expect(committed).toContain('MIT License');
		expect(committed).toContain('Copyright (c) Tailwind Labs, Inc.');
		expect(committed).toContain('The above copyright notice and this permission notice');
	});

	it('scopes every selector to a viewer root', () => {
		const selectors = selectorsOf(committed);
		expect(selectors.length).toBeGreaterThan(50);
		for (const selector of selectors) expect(selector).toContain(SCOPE);
	});

	it('keeps the specificity of every original selector', () => {
		// `:where()` adds no specificity, so removing the scope must give back a preflight selector
		// (or nothing, for the inherited defaults that now sit on the root itself).
		const original = new Set(selectorsOf(readInstalledPreflight().source));
		original.add('*');
		for (const selector of selectorsOf(committed)) {
			const bare = selector.replace(`:not(${SCOPE} *)`, '').replaceAll(SCOPE, '').trim();
			if (bare !== '') expect(original, selector).toContain(bare);
		}
	});

	it('starts the outermost root from the base text, whatever host text it sits in', () => {
		const rule = `${SCOPE}:not(${SCOPE} *) {\n\tfont-size: 1rem;\n\tfont-weight: normal;\n\tfont-style: normal;\n}`;
		expect(committed).toContain(rule);
	});

	it('leaves no page-level selector behind', () => {
		const stripped = committed.replace(/\/\*[\s\S]*?\*\//gu, '');
		expect(stripped).not.toMatch(/(^|[,}\s])(html|body|:host)\s*[,{]/u);
	});
});

describe('scopeSelector', () => {
	it('moves the html defaults onto the outermost root only', () => {
		expect(scopeSelector('html')).toStrictEqual([`${SCOPE}:not(${SCOPE} *)`]);
		expect(scopeSelector(':host')).toStrictEqual([`${SCOPE}:not(${SCOPE} *)`]);
	});

	it('drops the page body rule', () => {
		expect(scopeSelector('body')).toStrictEqual([]);
	});

	it('matches the root itself for universal and typeless selectors', () => {
		expect(scopeSelector('*')).toStrictEqual([SCOPE, `${SCOPE} *`]);
		expect(scopeSelector('::before')).toStrictEqual([`${SCOPE}::before`, `${SCOPE} ::before`]);
		expect(scopeSelector("[hidden]:where(:not([hidden='until-found']))")).toStrictEqual([
			`[hidden]:where(:not([hidden='until-found']))${SCOPE}`,
			`${SCOPE} [hidden]:where(:not([hidden='until-found']))`,
		]);
	});

	it('matches only descendants for type selectors', () => {
		expect(scopeSelector('h1')).toStrictEqual([`${SCOPE} h1`]);
		expect(scopeSelector(':where(select:is([multiple], [size])) optgroup')).toStrictEqual([
			`${SCOPE} :where(select:is([multiple], [size])) optgroup`,
		]);
	});
});
