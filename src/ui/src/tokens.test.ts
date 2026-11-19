// @vitest-environment jsdom
import { beforeAll, describe, expect, it } from 'vitest';
import { OFFICE_UI_TAGS, registerOfficeUi } from './index';
import { THEME_CSS } from './theme';
import { isDerivedToken, OFFICE_TOKENS, tok, type OfficeToken } from './tokens';

beforeAll(() => registerOfficeUi());

/** The CSS a control renders into its shadow root, as authored. */
function shadowCss(tag: string): string {
	const el = document.createElement(tag);
	document.body.append(el);
	const root = el.shadowRoot!;
	// The authored CSS: Lit keeps each style's source on the class. jsdom re-serialises adopted
	// rules (a bare `0` becomes `0px`), which would read as a raw length outside the tokens.
	const authored = (
		(el.constructor as { elementStyles?: ReadonlyArray<{ cssText?: string } | CSSStyleSheet> })
			.elementStyles ?? []
	).flatMap((style) =>
		'cssText' in style && typeof style.cssText === 'string' ? [style.cssText] : [],
	);
	const sheets = (root as ShadowRoot & { adoptedStyleSheets?: CSSStyleSheet[] }).adoptedStyleSheets;
	const text = authored.length
		? authored.join('\n')
		: sheets?.length
			? sheets.map((s) => [...s.cssRules].map((r) => r.cssText).join('\n')).join('\n')
			: [...root.querySelectorAll('style')].map((s) => s.textContent).join('\n');
	el.remove();
	return text;
}

/** Strip token references (with their defaults), media queries and keyframe percentages. */
function outsideTokens(css: string): string {
	const text = css.replace(/@media[^{]*\{/g, '{');
	// Remove each var(...) with its default, however deeply the default nests parentheses.
	let out = '';
	for (let i = 0; i < text.length;) {
		if (!text.startsWith('var(--', i)) {
			out += text[i++];
			continue;
		}
		let depth = 0;
		i += 'var'.length;
		do {
			if (text[i] === '(') depth++;
			else if (text[i] === ')') depth--;
			i++;
		} while (depth && i < text.length);
		out += 'T';
	}
	return out;
}

/** Set inline by the element that draws with them, so they have no theme default. */
const PER_ELEMENT = new Set([
	'--office-switch-knob-travel',
	'--office-gallery-preview-width',
	'--office-gallery-preview-height',
]);

describe('design tokens', () => {
	it('declares every literal token on :root in the theme, with the table default', () => {
		for (const [name, value] of Object.entries(OFFICE_TOKENS))
			if (!isDerivedToken(name as OfficeToken))
				expect(THEME_CSS, name).toContain(`${name}: ${value};`);
		expect(tok('--office-accent')).toBe('var(--office-accent, #2563eb)');
	});

	it('never declares a derived token in the theme, so it resolves where it is read', () => {
		// Declared on :root, `--office-accent-ink: var(--office-accent)` would be computed there, and
		// a container's own `--office-accent` would no longer reach the accent text inside it.
		const derived = Object.keys(OFFICE_TOKENS).filter((name) =>
			isDerivedToken(name as OfficeToken),
		);
		expect(derived).toContain('--office-accent-ink');
		expect(derived).toContain('--office-field-border');
		for (const name of derived) expect(THEME_CSS, name).not.toContain(`${name}:`);
		expect(tok('--office-field-border')).toBe(
			'var(--office-field-border, var(--office-control-border, #8a8886))',
		);
	});

	it('gives every control CSS free of raw colours and lengths outside tokens', () => {
		const offenders: string[] = [];
		for (const tag of OFFICE_UI_TAGS) {
			const raw = outsideTokens(shadowCss(tag)).match(
				// Negative lengths count too: a leading minus is part of the raw value.
				/#[0-9a-f]{3,8}\b|(?<![\w.])-?\d*\.?\d+px\b|rgba?\([^)]*\)/gi,
			);
			if (raw) offenders.push(`${tag}: ${[...new Set(raw)].join(', ')}`);
		}
		expect(offenders).toEqual([]);
	});

	it('references only declared tokens (private --_ helpers aside)', () => {
		const declared = new Set(Object.keys(OFFICE_TOKENS));
		const unknown = new Set<string>();
		for (const tag of OFFICE_UI_TAGS)
			for (const [, name] of shadowCss(tag).matchAll(/var\((--office-[\w-]+)/g))
				if (!declared.has(name!) && !PER_ELEMENT.has(name!)) unknown.add(name!);
		expect([...unknown]).toEqual([]);
	});

	it('lets a host retheme a control by overriding one token', () => {
		const button = document.createElement('office-ui-button');
		document.body.append(button);
		expect(shadowCss('office-ui-button')).toContain('var(--office-radius, 4px)');
		button.remove();
	});
});
