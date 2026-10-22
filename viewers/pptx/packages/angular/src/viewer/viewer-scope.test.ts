/**
 * The stylesheet's Tailwind preflight is scoped: it matches only inside
 * `[data-pptx-viewer]`, `[data-pptx-viewport]` and `[data-pptx-scope]`, so a
 * host page that imports `pptx-angular-viewer/styles` keeps its own headings,
 * lists and margins. The viewer host, and the canvas and ribbon a host can
 * render on their own, carry `data-pptx-scope` so everything they render
 * still gets the reset. This package has no TestBed, so the host bindings are
 * pinned as source text (the pattern `power-point-viewer-api.test.ts` uses).
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import {
	findScopedCustomElements,
	scopedCustomElementsIn,
} from '../../../../scripts/scope-marker-tags.mjs';
import { componentSource } from './component-source.test-support';

const HOST = "host: { 'data-pptx-scope': '' },";

describe('angular viewer roots carry the scoped reset attribute', () => {
	it.each([
		['power-point-viewer.component.ts', "selector: 'pptx-viewer'"],
		['slide-canvas.component.ts', "selector: 'pptx-slide-canvas'"],
		['ribbon.component.ts', "selector: 'pptx-ribbon'"],
	])('%s binds data-pptx-scope on its host', (file, selector) => {
		const source = componentSource(import.meta.dirname, file);
		const start = source.indexOf(selector);
		expect(start).toBeGreaterThan(-1);
		const decorator = source.slice(start, source.indexOf('template', start));
		expect(decorator).toContain(HOST);
		// An outermost scope takes the page's base text. These hosts are ordinary elements: with
		// emulated encapsulation their own `:host` rules become unlayered page rules, which beat
		// the layered reset. A shadow root would turn them into `:host` rules the reset overrides.
		expect(source).not.toContain('ViewEncapsulation.ShadowDom');
	});

	it('never puts a scope marker on a custom element in a template', () => {
		// A shared shadow-DOM control (`pptx-ui-context-menu`, ...) sets its own typography in
		// `:host`, which an outermost scope's base text would override: the marker goes on an
		// ordinary element around it. Component hosts are covered above.
		const src = path.resolve(import.meta.dirname, '..');
		expect(scopedCustomElementsIn(src, ['.html', '.ts'])).toStrictEqual([]);
		expect(
			findScopedCustomElements(
				'<pptx-ui-context-menu [attr.data-pptx-viewport]="on ? \'\' : null" />',
			),
		).toHaveLength(1);
		expect(
			findScopedCustomElements(
				'<div data-pptx-scope class="contents"><pptx-ui-context-menu /></div>',
			),
		).toStrictEqual([]);
	});

	it('scopes the compiled stylesheet to those roots instead of the whole page', () => {
		const css = readFileSync(
			path.resolve(import.meta.dirname, '../styles/pptx-angular-viewer.css'),
			'utf8',
		);
		expect(css).not.toMatch(/@import\s+'tailwindcss'\s*;/u);
		expect(css).toContain("@import 'tailwindcss/theme.css' layer(theme);");
		expect(css).toContain('src/ui/src/pptx/theme/viewer-preflight.css');
		expect(css).toContain("@import 'tailwindcss/utilities.css' layer(utilities);");
	});
});
