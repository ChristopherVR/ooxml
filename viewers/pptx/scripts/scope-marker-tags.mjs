/**
 * Finds the viewer's scope markers (`data-pptx-scope`, `data-pptx-viewer`,
 * `data-pptx-viewport`) on custom element tags in binding sources.
 *
 * The viewer stylesheets scope Tailwind's preflight to those markers, and an
 * outermost marked element also takes the page's base text (font size, weight,
 * style, line height and family) the way `<html>` did under the global reset.
 * On a shadow-DOM control that sets its own typography in `:host` (the
 * `pptx-ui-*` and `office-ui-*` elements), those document rules win over the
 * control's `:host` rules, so a menu teleported to `<body>` with the marker on
 * it grows from its own 12px to 16px. A marker therefore goes on an ordinary
 * element around the control (a `display: contents` wrapper when nothing else
 * fits), never on a tag with a hyphen. Each binding's tests run this over its
 * sources.
 *
 * The scan reads markup the way the three template syntaxes write it: Vue
 * templates (`data-pptx-scope`, `:data-pptx-scope`, `v-bind:data-pptx-scope`),
 * JSX (`data-pptx-scope=''`) and Angular templates
 * (`[attr.data-pptx-viewport]`). Attribute values are skipped whole, so a
 * selector string such as `closest('[data-pptx-viewport]')` is not a marker.
 * Angular component `host:` bindings are not tags and are checked by the
 * Angular tests.
 */
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

/** The attributes the scoped preflight matches. */
export const SCOPE_MARKERS = ['data-pptx-scope', 'data-pptx-viewer', 'data-pptx-viewport'];

// An opening tag: the name, then attributes where quoted values (which may hold `>`) and one
// level of nested JSX braces (`style={{ a: 1 }}`) are consumed whole.
const OPENING_TAG =
	/<([A-Za-z][\w.:-]*)((?:[^>"'{}]|"[^"]*"|'[^']*'|\{(?:[^{}]|\{[^{}]*\})*\})*?)\/?>/gu;
const QUOTED_VALUE = /=\s*(?:"[^"]*"|'[^']*'|\{(?:[^{}]|\{[^{}]*\})*\})/gu;
const MARKER = new RegExp(
	`(?:^|\\s)(?:\\[attr\\.|:|v-bind:)?(${SCOPE_MARKERS.join('|')})(?=[\\s=\\]/>]|$)`,
	'u',
);

/** A custom element tag name: lower-case start and a hyphen (`pptx-ui-context-menu`). */
export function isCustomElementTag(name) {
	return /^[a-z][\w.]*-[\w.-]*$/u.test(name);
}

/**
 * Every opening tag in `source` whose name is a custom element and that carries a scope marker.
 *
 * @param {string} source
 * @returns {{ tag: string, marker: string, line: number }[]}
 */
export function findScopedCustomElements(source) {
	const found = [];
	for (const match of source.matchAll(OPENING_TAG)) {
		const [whole, name, attributes] = match;
		if (!isCustomElementTag(name)) {
			continue;
		}
		const marker = MARKER.exec(attributes.replace(QUOTED_VALUE, '=""'));
		if (marker) {
			const line = source.slice(0, match.index).split('\n').length;
			found.push({ tag: whole.replace(/\s+/gu, ' '), marker: marker[1], line });
		}
	}
	return found;
}

/**
 * The non-test source files under `dir` with one of `extensions`.
 *
 * @param {string} dir
 * @param {readonly string[]} extensions
 * @returns {string[]}
 */
export function sourceFiles(dir, extensions) {
	return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
		const file = join(dir, entry.name);
		if (entry.isDirectory()) {
			return entry.name === 'node_modules' || entry.name === 'dist'
				? []
				: sourceFiles(file, extensions);
		}
		const isTest = /\.(?:test|spec)\.[cm]?[jt]sx?$/u.test(entry.name);
		return !isTest && extensions.some((extension) => entry.name.endsWith(extension)) ? [file] : [];
	});
}

/**
 * `file:line <tag>` for every scope marker on a custom element under `dir`.
 *
 * @param {string} dir
 * @param {readonly string[]} extensions
 * @returns {string[]}
 */
export function scopedCustomElementsIn(dir, extensions) {
	return sourceFiles(dir, extensions).flatMap((file) =>
		findScopedCustomElements(readFileSync(file, 'utf8')).map(
			({ tag, line }) => `${file}:${line} ${tag}`,
		),
	);
}
