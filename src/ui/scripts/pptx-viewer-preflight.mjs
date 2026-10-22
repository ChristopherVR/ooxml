// Generates `src/pptx/theme/viewer-preflight.css`: Tailwind's preflight (its base reset) with
// every selector scoped to the PowerPoint viewer, so a host page that loads a viewer stylesheet
// keeps its own headings, lists, margins and images. The React, Vue and Angular stylesheets import
// the generated file in place of the global preflight that `@import 'tailwindcss'` brings in.
//
//   node scripts/pptx-viewer-preflight.mjs          rewrites the committed file
//   node scripts/pptx-viewer-preflight.mjs --check  fails when it is out of date
//
// `src/pptx/theme/viewer-preflight.test.ts` fails when the committed file differs from the
// transform of the installed `tailwindcss/preflight.css`, so a Tailwind upgrade cannot silently
// leave the viewer on an old reset.
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * The elements a viewer stylesheet styles: the React viewer root, any canvas viewport (also a
 * standalone `SlideCanvas`), and `data-pptx-scope`, which every other root carries: the Vue and
 * Angular viewer roots, standalone building blocks, and surfaces a binding renders outside its
 * root (dialogs and menus teleported to `document.body`).
 */
export const VIEWER_SCOPES = ['[data-pptx-viewer]', '[data-pptx-viewport]', '[data-pptx-scope]'];

export const OUTPUT = fileURLToPath(
	new URL('../src/pptx/theme/viewer-preflight.css', import.meta.url),
);

/** Splits `text` at top-level occurrences of `separator`, ignoring parentheses and strings. */
function splitTopLevel(text, separator) {
	const parts = [];
	let depth = 0;
	let quote = '';
	let start = 0;
	for (let index = 0; index < text.length; index++) {
		const char = text[index];
		if (quote) {
			if (char === quote) quote = '';
		} else if (char === '"' || char === "'") {
			quote = char;
		} else if (char === '(' || char === '[') {
			depth++;
		} else if (char === ')' || char === ']') {
			depth--;
		} else if (depth === 0 && char === separator) {
			parts.push(text.slice(start, index));
			start = index + 1;
		}
	}
	parts.push(text.slice(start));
	return parts.map((part) => part.trim()).filter(Boolean);
}

/** True when `selector` is one compound selector with no type selector (it could match a root). */
function isTypelessCompound(selector) {
	return /^[[:]/u.test(selector) && splitTopLevel(selector, ' ').length === 1;
}

/** Index of the first top-level pseudo-element (`::x`) in a compound selector, or -1. */
function pseudoElementIndex(selector) {
	let depth = 0;
	for (let index = 0; index < selector.length - 1; index++) {
		const char = selector[index];
		if (char === '(' || char === '[') depth++;
		else if (char === ')' || char === ']') depth--;
		else if (depth === 0 && char === ':' && selector[index + 1] === ':') return index;
	}
	return -1;
}

/**
 * Rewrites one preflight selector to match only inside the viewer. Every result keeps the
 * specificity of the original, because `:where()` adds none, so the viewer's own rules win and
 * lose exactly as they did against the global reset.
 *
 * - `html` and `:host` (the inherited defaults: line height, font, tab size) apply to the outermost
 *   root, never to a root nested in another one, which keeps inheriting from its viewer.
 * - `body` styles the page, so it is dropped.
 * - `*`, a bare pseudo-element (`::before` is `*::before`) and any other selector without a type
 *   (`[hidden]`) match the root itself as well as everything inside it.
 * - Everything else matches descendants of a root.
 */
export function scopeSelector(selector, scopes = VIEWER_SCOPES) {
	const scope = `:where(${scopes.join(',')})`;
	if (selector === 'html' || selector === ':host') return [`${scope}:not(${scope} *)`];
	if (selector === 'body') return [];
	if (selector === '*') return [scope, `${scope} *`];
	if (selector.startsWith('::')) return [`${scope}${selector}`, `${scope} ${selector}`];
	if (isTypelessCompound(selector)) {
		const pseudo = pseudoElementIndex(selector);
		const self =
			pseudo === -1
				? `${selector}${scope}`
				: `${selector.slice(0, pseudo)}${scope}${selector.slice(pseudo)}`;
		return [self, `${scope} ${selector}`];
	}
	return [`${scope} ${selector}`];
}

/** Parses a stylesheet into comments, qualified rules and block at-rules (enough for preflight). */
function parse(css) {
	let index = 0;
	function nodes() {
		const out = [];
		while (index < css.length) {
			const rest = css.slice(index);
			const space = /^\s+/u.exec(rest);
			if (space) {
				index += space[0].length;
				continue;
			}
			if (rest.startsWith('/*')) {
				const end = css.indexOf('*/', index + 2);
				out.push({ type: 'comment', text: css.slice(index, end + 2) });
				index = end + 2;
				continue;
			}
			if (rest.startsWith('}')) {
				index++;
				return out;
			}
			const open = css.indexOf('{', index);
			if (open === -1) throw new Error(`Unexpected text in preflight: ${rest.slice(0, 40)}`);
			// Comments inside a prelude (Tailwind annotates its `@supports` condition) are kept.
			const prelude = css.slice(index, open).trim();
			index = open + 1;
			if (prelude.startsWith('@')) {
				out.push({ type: 'at-rule', prelude, children: nodes() });
				continue;
			}
			const close = css.indexOf('}', index);
			out.push({ type: 'rule', prelude, body: css.slice(index, close) });
			index = close + 1;
		}
		return out;
	}
	return nodes();
}

/** Re-indents a declaration block, one declaration per line, keeping its comments and values. */
function printBody(body, indent) {
	const lines = body
		.split('\n')
		.map((line) => line.trim())
		.filter(Boolean);
	// A multi-line value (`font-family: --theme(...)`) is joined back onto one line.
	const declarations = [];
	let current = '';
	let depth = 0;
	for (const line of lines) {
		const gap = current.endsWith('(') || line.startsWith(')') ? '' : ' ';
		current = current ? `${current}${gap}${line}` : line;
		for (const char of line) {
			if (char === '(') depth++;
			else if (char === ')') depth--;
		}
		if (depth === 0) {
			declarations.push(current);
			current = '';
		}
	}
	return declarations.map((line) => `${indent}\t${line}`).join('\n');
}

function print(nodes, indent = '') {
	const blocks = [];
	for (const node of nodes) {
		if (node.type === 'comment') {
			blocks.push(
				node.text
					.split('\n')
					.map((line, at) => (at === 0 ? `${indent}${line}` : `${indent}${line.trimEnd()}`))
					.join('\n'),
			);
			continue;
		}
		if (node.type === 'at-rule') {
			const prelude = node.prelude.replace(/\s*\n\s*/gu, ' ');
			blocks.push(`${indent}${prelude} {\n${print(node.children, `${indent}\t`)}\n${indent}}`);
			continue;
		}
		const selectors = [
			...new Set(splitTopLevel(node.prelude, ',').flatMap((selector) => scopeSelector(selector))),
		];
		if (selectors.length === 0) continue;
		const head = selectors.map((selector) => `${indent}${selector}`).join(',\n');
		blocks.push(`${head} {\n${printBody(node.body, indent)}\n${indent}}`);
	}
	// A comment stays attached to the rule it documents, as in the source.
	return blocks
		.map((block, at) =>
			at > 0 && !blocks[at - 1].trimStart().startsWith('/*') ? `\n${block}` : block,
		)
		.join('\n');
}

/**
 * Not part of Tailwind's preflight. An outermost root can sit inside host text the browser
 * styles, such as a heading or `<small>`, which the global reset flattened to the page's base
 * text; scoped, the reset leaves that host text alone, so the root starts from the base text
 * itself. Only the outermost root: a nested root keeps inheriting from its viewer.
 */
export function rootTextRule(scopes = VIEWER_SCOPES) {
	const scope = `:where(${scopes.join(',')})`;
	return [
		'/*',
		'  Not from Tailwind: start the outermost viewer root from the base text, whatever host text',
		'  it sits in (a heading, `<small>`). The global reset did this by flattening that host text.',
		'*/',
		`${scope}:not(${scope} *) {`,
		'\tfont-size: 1rem;',
		'\tfont-weight: normal;',
		'\tfont-style: normal;',
		'}',
	].join('\n');
}

/** Builds the scoped stylesheet from Tailwind's `preflight.css` text. */
export function scopePreflight(source, { version, license }) {
	const notice = license
		.trim()
		.split('\n')
		.map((line) => (line ? ` * ${line}` : ' *'))
		.join('\n');
	const header = [
		'/*',
		` * Generated by src/ui/scripts/pptx-viewer-preflight.mjs from tailwindcss@${version}`,
		' * preflight.css. Do not edit: run `node scripts/pptx-viewer-preflight.mjs` in src/ui.',
		' *',
		' * Tailwind preflight scoped to the PowerPoint viewer: every rule matches only inside',
		` * ${VIEWER_SCOPES.join(', ')},`,
		' * through zero-specificity `:where()`, so a host page keeps its own base styles. Import it',
		' * into the `base` layer of a viewer stylesheet in place of `tailwindcss/preflight.css`.',
		' *',
		' * Tailwind CSS is distributed under the MIT License:',
		' *',
		notice,
		' */',
	].join('\n');
	return `${header}\n\n${print(parse(source))}\n\n${rootTextRule()}\n`;
}

/** Reads the installed preflight and its version and licence. */
export function readInstalledPreflight() {
	const require = createRequire(import.meta.url);
	const file = require.resolve('tailwindcss/preflight.css');
	const root = dirname(file);
	return {
		source: readFileSync(file, 'utf8'),
		version: JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')).version,
		license: readFileSync(join(root, 'LICENSE'), 'utf8'),
	};
}

export function generate() {
	const { source, version, license } = readInstalledPreflight();
	return scopePreflight(source, { version, license });
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
	const css = generate();
	if (process.argv.includes('--check')) {
		if (readFileSync(OUTPUT, 'utf8') !== css) {
			console.error(`${OUTPUT} is out of date: run node scripts/pptx-viewer-preflight.mjs`);
			process.exit(1);
		}
	} else {
		writeFileSync(OUTPUT, css);
	}
}
