// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';

import { EDITOR_CHROME_CSS } from './index';
import { EDITOR_CARDS_CSS } from './cards-css';

afterEach(() => {
	document.body.replaceChildren();
	document.documentElement.removeAttribute('data-office-look');
});

const SHELL = `<div data-pptx-chrome="body">
	<aside data-pptx-chrome="slides"></aside>
	<div role="separator"></div>
	<main><div data-pptx-viewport></div></main>
	<div><div data-pptx-chrome="inspector"><div data-pptx-chrome="inspector-tabs"><button aria-selected="true"></button></div></div></div>
</div><div data-pptx-chrome="notes"></div>`;

function mount(lookOn: 'none' | 'self' | 'ancestor') {
	const style = document.createElement('style');
	style.textContent = EDITOR_CARDS_CSS;
	const host = document.createElement('div');
	host.dataset.pptxEditorChrome = '';
	host.innerHTML = SHELL;
	if (lookOn === 'self') host.setAttribute('data-office-look', 'cards');
	if (lookOn === 'ancestor') document.documentElement.setAttribute('data-office-look', 'cards');
	document.body.append(style, host);
	const rules = [...style.sheet!.cssRules]
		.filter((rule): rule is CSSMediaRule => rule instanceof CSSMediaRule)
		.filter((rule) => rule.conditionText === '(min-width: 768px)')
		.flatMap((rule) => [...rule.cssRules] as CSSStyleRule[]);
	const value = (selector: string, property: string): string | undefined =>
		rules
			.filter((rule) => !rule.selectorText.includes('::'))
			.filter((rule) => host.querySelector(selector)?.matches(rule.selectorText))
			.map((rule) => rule.style.getPropertyValue(property))
			.filter(Boolean)
			.at(-1);
	return { host, rules, value };
}

describe('cards look stylesheet', () => {
	it('changes nothing without the look', () => {
		const { value } = mount('none');
		expect(value('[data-pptx-chrome="slides"]', 'border-radius')).toBeUndefined();
		expect(value('[data-pptx-viewport]', 'box-shadow')).toBeUndefined();
	});

	it.each(['self', 'ancestor'] as const)('floats the chrome as cards with the look on %s', (on) => {
		const { value } = mount(on);
		for (const region of ['slides', 'inspector', 'notes'])
			expect(value(`[data-pptx-chrome="${region}"]`, 'border-radius')).toBe(
				'var(--office-radius-surface, 14px)',
			);
		expect(value('[data-pptx-viewport]', 'border-radius')).toBe(
			'var(--office-radius-surface, 14px)',
		);
		expect(value('[data-pptx-chrome="slides"]', 'box-shadow')).toBe('var(--_pptx-card-shadow)');
		expect(value('[data-pptx-chrome="body"]', 'gap')).toBe('var(--_pptx-gap)');
		expect(value('[role="separator"]', 'margin-inline')).toBe('calc(-1 * var(--_pptx-gap))');
		expect(value('[data-pptx-chrome="inspector"]', 'min-width')).toBe(
			'var(--office-side-pane-width, 320px)',
		);
		expect(value('[data-pptx-chrome="inspector-tabs"] button', 'background')).toBe(
			'var(--_pptx-card)',
		);
	});

	it('keeps the palette: no literal colours outside the shadow fallback', () => {
		const withoutShadow = EDITOR_CARDS_CSS.replaceAll(/rgb\(60 64 67 \/ \d+%\)/g, '');
		expect(withoutShadow).not.toMatch(/#[0-9a-f]{3,8}\b|rgba?\(/i);
	});

	it('is the last sheet of the chrome, so it wins at equal weight', () => {
		expect(EDITOR_CHROME_CSS.endsWith(EDITOR_CARDS_CSS)).toBe(true);
	});
});
