import { describe, expect, it } from 'vitest';

import { buildViewerCss } from './css';

/** The rules inside the viewer's `@media (forced-colors: active)` blocks, as written. */
function forcedRules(): { selector: string; body: string }[] {
	const css = buildViewerCss().replace(/\/\*[\s\S]*?\*\//g, '');
	return [...css.matchAll(/@media[^{]*forced-colors:\s*active[^{]*\{/g)].flatMap((match) => {
		let depth = 1;
		let end = match.index + match[0].length;
		while (depth && end < css.length) {
			const char = css[end++];
			if (char === '{') depth++;
			else if (char === '}') depth--;
		}
		const block = css.slice(match.index + match[0].length, end - 1);
		return [...block.matchAll(/([^{}]+)\{([^{}]*)\}/g)].map((rule) => ({
			selector: rule[1]!.trim(),
			body: rule[2]!,
		}));
	});
}

/** The forced-colours declarations that reach `el`, normalised to `property: value`. */
function forcedPaint(el: Element): string[] {
	return forcedRules()
		.filter((rule) => {
			try {
				return el.matches(rule.selector);
			} catch {
				// The test DOM cannot parse some focus selectors; they paint no selected state.
				return false;
			}
		})
		.flatMap((rule) => rule.body.split(';').map((part) => part.trim().replace(/\s+/g, ' ')))
		.filter(Boolean);
}

describe('selected states in forced colours', () => {
	const root = document.createElement('div');
	root.className = 'pptxv';
	root.innerHTML = `
		<button class="pptxv-thumb is-active" aria-current="true"></button>
		<button class="pptxv-thumb"></button>
		<div class="pptxv-slide-size-orientation"><button class="is-active" aria-pressed="true">Landscape</button></div>
		<button class="pptxv-btn is-active" aria-pressed="true"></button>
		<button class="pptxv-inspector-tab is-active" role="tab" aria-selected="true">Properties</button>
		<button class="pptxv-inspector-tab" role="tab" aria-selected="false">Animations</button>`;
	document.body.append(root);
	const [current, other, landscape, toggle, tab, idleTab] = [...root.querySelectorAll('button')];

	it('outlines the current slide in the system highlight, and no other slide', () => {
		expect(forcedPaint(current!)).toContain('outline: 2px solid Highlight !important');
		expect(forcedPaint(other!).join()).not.toContain('Highlight');
	});

	it('fills pressed toggles and the selected tab with the highlight, over any accent rule', () => {
		for (const el of [landscape!, toggle!, tab!]) {
			const paint = forcedPaint(el);
			expect(paint).toContain('background: Highlight !important');
			expect(paint).toContain('color: HighlightText !important');
			expect(paint).toContain('forced-color-adjust: none');
		}
		expect(forcedPaint(idleTab!).join()).not.toContain('Highlight !important');
	});
});
