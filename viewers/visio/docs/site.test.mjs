import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';
import postcss from 'postcss';
import { initVisioTheme } from './assets/theme.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const pages = ['index.html', 'docs/index.html', 'docs/parity.html', 'docs/architecture.html'];
const read = (path) => readFileSync(resolve(root, path), 'utf8');
const dom = (path) => new JSDOM(read(path), { url: `https://example.test/visio-viewer/${path}` });
const css = postcss.parse(read('docs/assets/site.css'));

const normalizeCss = (value) => value.replace(/\s+/g, '').replace(/["']/g, '');
function rule(selector) {
	return css.nodes.find(
		(node) => node.type === 'rule' && normalizeCss(node.selector) === normalizeCss(selector),
	);
}
function declarations(node) {
	return Object.fromEntries(
		node.nodes.filter((n) => n.type === 'decl').map((n) => [n.prop, n.value]),
	);
}
function contrast(a, b) {
	function luminance(hex) {
		const rgb = [1, 3, 5].map((start) => parseInt(hex.slice(start, start + 2), 16) / 255);
		const linear = rgb.map((c) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
		return linear.reduce((sum, c, i) => sum + c * [0.2126, 0.7152, 0.0722][i], 0);
	}
	const values = [luminance(a), luminance(b)].sort((x, y) => x - y);
	return (values[1] + 0.05) / (values[0] + 0.05);
}

describe('static documentation contracts', () => {
	for (const page of pages) {
		it(`${page}: semantic landmarks, heading, IDs and image alternatives`, () => {
			const { document } = dom(page).window;
			assert.equal(document.documentElement.lang, 'en');
			assert.equal(document.querySelectorAll('main').length, 1);
			assert.equal(document.querySelectorAll('h1').length, 1);
			assert.ok(document.querySelector('a.skip[href="#main"]'));
			const ids = [...document.querySelectorAll('[id]')].map((n) => n.id);
			assert.equal(ids.length, new Set(ids).size, 'IDs must be unique');
			for (const image of document.images) assert.ok(image.hasAttribute('alt'));
			for (const button of document.querySelectorAll('button')) {
				assert.equal(button.type, 'button');
				assert.ok(button.getAttribute('aria-label') || button.textContent.trim());
			}
		});
		it(`${page}: relative local links resolve under a repository prefix`, () => {
			const { document } = dom(page).window;
			for (const element of document.querySelectorAll('[href], [src], [data-src]')) {
				const value =
					element.getAttribute('href') ??
					element.getAttribute('src') ??
					element.getAttribute('data-src');
				if (/^(https?:|mailto:|#)/.test(value)) continue;
				assert.ok(!value.startsWith('/'), `Root-absolute URL breaks repository hosting: ${value}`);
				const target = resolve(root, dirname(page), decodeURIComponent(value.split(/[?#]/)[0]));
				assert.ok(existsSync(target), `Missing local target: ${value}`);
			}
		});
		it(`${page}: no third-party scripts, styles or image assets`, () => {
			const { document } = dom(page).window;
			for (const element of document.querySelectorAll(
				'script[src], link[rel="stylesheet"], img[src], iframe[src]',
			)) {
				const url = element.getAttribute('src') ?? element.getAttribute('href');
				assert.ok(!/^(https?:)?\/\//.test(url), `Unexpected remote asset: ${url}`);
			}
			for (const region of document.querySelectorAll('.table-scroll')) {
				assert.equal(region.getAttribute('tabindex'), '0');
				assert.equal(region.getAttribute('role'), 'region');
				assert.ok(region.getAttribute('aria-label'));
			}
		});
	}
	it('uses readable contrast for body, secondary, accent and button text', () => {
		for (const selector of [':root', ':root[data-theme=dark]']) {
			const vars = declarations(rule(selector));
			for (const name of ['--ink', '--muted', '--accent']) {
				const ratio = contrast(vars[name], vars['--paper']);
				assert.ok(ratio >= 4.5, `${selector} ${name}: ${ratio.toFixed(2)}:1`);
			}
			const buttonText = selector === ':root' ? '#ffffff' : vars['--paper'];
			assert.ok(contrast(buttonText, vars['--accent']) >= 4.5);
		}
	});
	it('preserves explicit mobile and reduced-motion source rules', () => {
		const media = css.nodes.filter((node) => node.type === 'atrule' && node.name === 'media');
		const mobile = media.find((node) => normalizeCss(node.params) === '(max-width:560px)');
		const tablet = media.find((node) => normalizeCss(node.params) === '(max-width:820px)');
		const motion = media.find(
			(node) => normalizeCss(node.params) === '(prefers-reduced-motion:reduce)',
		);
		assert.ok(mobile && tablet && motion);
		const has = (parent, selector, prop, value) =>
			parent.nodes.some(
				(n) =>
					n.type === 'rule' &&
					normalizeCss(n.selector) === normalizeCss(selector) &&
					declarations(n)[prop] === value,
			);
		assert.ok(has(mobile, '.feature-grid', 'grid-template-columns', '1fr'));
		assert.ok(has(tablet, '.hero-grid', 'grid-template-columns', '1fr'));
		assert.ok(has(tablet, '.doc-layout', 'display', 'block'));
		assert.equal(declarations(rule('.table-scroll')).overflow, 'auto');
		assert.equal(declarations(rule('.code-section>div,.code-card'))['min-width'], '0');
		assert.equal(
			declarations(rule('.theme-toggle'))['min-height'] ??
				declarations(rule('.theme-toggle')).height,
			'44px',
		);
		assert.equal(declarations(rule('.frameworks button'))['min-height'], '44px');
		assert.ok(
			motion.nodes.some(
				(n) =>
					n.type === 'rule' &&
					n.nodes.some(
						(d) => d.type === 'decl' && d.prop === 'animation' && d.value === 'none' && d.important,
					),
			),
		);
		let hiddenContent = false;
		css.walkDecls('opacity', (decl) => {
			if (decl.value === '0') hiddenContent = true;
		});
		assert.ok(!hiddenContent, 'Content must not be hidden pending JavaScript');
	});
	it('keeps sample and fidelity limitations explicit', () => {
		const home = dom('index.html').window.document.body.textContent.replace(/\s+/g, ' ');
		const ledger = dom('docs/parity.html').window.document.body.textContent;
		assert.match(home, /Original illustration · not a live editor/);
		assert.match(home, /public beta/i);
		assert.match(home, /npm packages are placeholders without a viewer API/i);
		assert.match(home, /General drawing/);
		assert.match(home, /Experimental source-backed plain-text/);
		assert.match(home, /Native Visio reopening remains unverified/);
		assert.doesNotMatch(home, /private and unpublished|npm install|MIT license/i);
		assert.match(ledger, /parity is a target, not the current result/i);
		assert.match(ledger, /Generated fixtures/);
		const rows = [...dom('docs/parity.html').window.document.querySelectorAll('tbody tr')];
		assert.ok(rows.length >= 29, 'The ledger must retain the established capability inventory');
		const names = rows.map((row) => row.cells[0].textContent.trim());
		assert.equal(names.length, new Set(names).size, 'Capability names must be unique');
		for (const required of ['ShapeSheet formulas', 'Corner rounding', 'Visio visual parity'])
			assert.ok(names.includes(required), `Missing capability: ${required}`);
	});
});

describe('documentation interaction logic in a simulated DOM', () => {
	function interactive(configure = () => {}) {
		const instance = new JSDOM(read('index.html'), {
			url: 'https://example.test/visio-viewer/',
			runScripts: 'outside-only',
		});
		instance.window.matchMedia = () => ({ matches: false });
		configure(instance.window);
		initVisioTheme(instance.window);
		instance.window.eval(
			read('docs/assets/site.js')
				.replace(/^import .*;\r?\n/m, '')
				.replace('initVisioTheme(window);', ''),
		);
		return instance;
	}
	it('native framework tabs select one labelled local-source example', () => {
		const instance = interactive();
		const document = instance.window.document;
		const choices = [...document.querySelectorAll('[data-framework]')];
		assert.deepEqual(choices.map((choice) => choice.dataset.framework).sort(), [
			'angular',
			'react',
			'solid',
			'svelte',
			'vanilla',
			'vue',
		]);
		assert.equal(document.querySelectorAll('[role="tab"]').length, 6);
		for (const choice of choices) {
			choice.click();
			assert.equal(document.querySelectorAll('[data-framework][aria-selected="true"]').length, 1);
			assert.equal(choice.tabIndex, 0);
			assert.equal(
				document.querySelector('#integration-example').getAttribute('aria-labelledby'),
				choice.id,
			);
			const content = document.querySelector('#integration-code');
			if (choice.dataset.framework === 'vanilla') {
				assert.match(content.textContent, /mountViewer/);
				assert.match(content.textContent, /destroy/);
			} else {
				assert.match(content.textContent, /VisioViewer/);
				assert.match(content.textContent, /\.\/packages\/bindings\/src\//);
			}
			assert.match(content.textContent, /\.\/src\/index/);
			assert.equal(content.children.length, 0, 'Code is text, never injected HTML');
		}
		instance.window.close();
	});
	it('code tabs use roving keyboard focus with Arrow, Home, and End', () => {
		const instance = interactive();
		const { document, KeyboardEvent } = instance.window;
		const choices = [...document.querySelectorAll('[data-framework]')];
		const press = (tab, key) =>
			tab.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true }));
		choices[0].focus();
		press(choices[0], 'ArrowLeft');
		assert.equal(document.activeElement, choices.at(-1));
		press(choices.at(-1), 'ArrowRight');
		assert.equal(document.activeElement, choices[0]);
		press(choices[0], 'End');
		assert.equal(document.activeElement, choices.at(-1));
		press(choices.at(-1), 'Home');
		assert.equal(document.activeElement, choices[0]);
		assert.equal(choices.filter((choice) => choice.tabIndex === 0).length, 1);
		instance.window.close();
	});
	it('mobile menu toggles, closes on Escape, and returns focus', () => {
		const instance = interactive();
		const { document, KeyboardEvent } = instance.window;
		const menu = document.querySelector('.menu-toggle');
		const navigation = document.getElementById(menu.getAttribute('aria-controls'));
		menu.click();
		assert.equal(menu.getAttribute('aria-expanded'), 'true');
		assert.equal(navigation.dataset.open, 'true');
		navigation.querySelector('a').focus();
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		assert.equal(menu.getAttribute('aria-expanded'), 'false');
		assert.equal(document.activeElement, menu);
		menu.click();
		menu.click();
		assert.equal(navigation.dataset.open, 'false');
		instance.window.close();
	});
	it('theme toggle updates shared preference, frame, metadata, and accessible label', () => {
		const instance = interactive();
		const { document, localStorage } = instance.window;
		const button = document.querySelector('.theme-toggle');
		button.click();
		assert.equal(document.documentElement.dataset.theme, 'dark');
		assert.equal(localStorage.getItem('visio-docs-theme'), 'dark');
		assert.equal(button.getAttribute('aria-label'), 'Switch to light theme');
		assert.equal(document.querySelector('meta[name="theme-color"]').content, '#0f1113');
		assert.equal(
			document.getElementById('live-viewer').contentDocument.documentElement.dataset.theme,
			'dark',
		);
		button.click();
		assert.equal(document.documentElement.dataset.theme, 'light');
		instance.window.close();
	});
	it('saved theme and later workspace storage changes synchronize the page', () => {
		const instance = interactive((window) =>
			window.localStorage.setItem('visio-docs-theme', 'dark'),
		);
		const { document, StorageEvent } = instance.window;
		assert.equal(document.documentElement.dataset.theme, 'dark');
		instance.window.dispatchEvent(
			new StorageEvent('storage', { key: 'visio-docs-theme', newValue: 'light' }),
		);
		assert.equal(document.documentElement.dataset.theme, 'light');
		instance.window.dispatchEvent(
			new StorageEvent('storage', { key: 'unrelated', newValue: 'dark' }),
		);
		assert.equal(document.documentElement.dataset.theme, 'light');
		instance.window.close();
	});
	it('system changes follow until the visitor chooses a theme, even without storage', () => {
		let systemChanged;
		const media = {
			matches: true,
			addEventListener: (_, listener) => {
				systemChanged = listener;
			},
		};
		const instance = interactive((window) => {
			window.matchMedia = () => media;
			Object.defineProperty(window, 'localStorage', {
				get() {
					throw new Error('Unavailable');
				},
			});
		});
		const { document } = instance.window;
		assert.equal(document.documentElement.dataset.theme, 'dark');
		media.matches = false;
		systemChanged();
		assert.equal(document.documentElement.dataset.theme, 'light');
		document.querySelector('.theme-toggle').click();
		systemChanged();
		assert.equal(document.documentElement.dataset.theme, 'dark');
		instance.window.close();
	});
	it('unavailable clipboard gives an announced recovery message', async () => {
		const instance = interactive();
		instance.window.document.querySelector('.copy-code').click();
		await Promise.resolve();
		const status = instance.window.document.querySelector('#copy-status');
		assert.equal(status.getAttribute('role'), 'status');
		assert.match(status.textContent, /Select and copy/);
		instance.window.close();
	});
	it('late clipboard completion cannot announce copying a newly selected example', async () => {
		let finish;
		const instance = interactive((window) => {
			Object.defineProperty(window.navigator, 'clipboard', {
				value: {
					writeText: () =>
						new Promise((resolve) => {
							finish = resolve;
						}),
				},
			});
		});
		instance.window.document.querySelector('.copy-code').click();
		instance.window.document.querySelector('[data-framework="react"]').click();
		finish();
		await Promise.resolve();
		assert.equal(instance.window.document.querySelector('#copy-status').textContent, '');
		instance.window.close();
	});
	it('lazy live viewer admits one load and only the expected readiness sender', () => {
		let timeout;
		const instance = interactive((window) => {
			window.setTimeout = (callback) => {
				timeout = callback;
				return 1;
			};
		});
		instance.window.eval(read('docs/assets/live-demo.js'));
		const { document, MessageEvent } = instance.window;
		const frame = document.getElementById('live-viewer');
		const button = document.getElementById('load-demo');
		const status = document.getElementById('demo-status');
		assert.equal(frame.hasAttribute('src'), false);
		button.click();
		assert.equal(frame.getAttribute('src'), 'demo/index.html?embed=1');
		assert.equal(frame.hidden, false);
		const source = frame.contentWindow;
		// jsdom does not fetch the iframe HTML; supply the ready document root.
		frame.contentDocument.appendChild(frame.contentDocument.createElement('html'));
		button.click();
		assert.equal(frame.contentWindow, source, 'Repeated clicks must not reload a live document');
		const ready = (origin, sender, data = { type: 'visio-viewer-ready' }) =>
			instance.window.dispatchEvent(new MessageEvent('message', { origin, source: sender, data }));
		ready('https://untrusted.test', source);
		ready(instance.window.location.origin, instance.window);
		ready(instance.window.location.origin, source, { type: 'other' });
		assert.match(status.textContent, /Loading/);
		timeout();
		assert.match(status.textContent, /taking longer/);
		ready(instance.window.location.origin, source);
		assert.match(status.textContent, /Viewer ready/);
		assert.equal(frame.getAttribute('aria-busy'), 'false');
		assert.equal(
			frame.contentDocument.documentElement.dataset.theme,
			document.documentElement.dataset.theme,
		);
		timeout();
		assert.match(status.textContent, /Viewer ready/);
		instance.window.close();
	});
});
