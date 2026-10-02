import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';
import postcss from 'postcss';

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
			for (const element of document.querySelectorAll('[href], [src]')) {
				const value = element.getAttribute('href') ?? element.getAttribute('src');
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
		const home = dom('index.html').window.document.body.textContent;
		const ledger = dom('docs/parity.html').window.document.body.textContent;
		assert.match(home, /Illustration · not a live editor/);
		assert.match(home, /public beta/i);
		assert.match(home, /npm packages are placeholders without a viewer API/i);
		assert.match(home, /General drawing edits/);
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
	function interactive() {
		const instance = new JSDOM(read('index.html'), {
			url: 'https://example.test/visio-viewer/',
			runScripts: 'outside-only',
		});
		instance.window.matchMedia = () => ({ matches: false });
		instance.window.eval(read('docs/assets/site.js'));
		return instance;
	}
	it('each framework button selects exactly one labelled code example', () => {
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
		assert.equal(
			document.querySelectorAll('[role="tab"]').length,
			0,
			'Use native toggle buttons, not incomplete ARIA tabs',
		);
		for (const choice of choices) {
			choice.click();
			assert.equal(document.querySelectorAll('[data-framework][aria-pressed="true"]').length, 1);
			assert.equal(choice.getAttribute('aria-pressed'), 'true');
			assert.match(document.querySelector('#integration-code').textContent, /mountViewer/);
			assert.match(document.querySelector('#integration-code').textContent, /destroy/);
			assert.equal(
				document.querySelector('#integration-code').children.length,
				0,
				'Code is text, never injected HTML',
			);
		}
		instance.window.close();
	});
	it('theme toggle updates preference and accessible label', () => {
		const instance = interactive();
		const { document, localStorage } = instance.window;
		const button = document.querySelector('.theme-toggle');
		button.click();
		assert.equal(document.documentElement.dataset.theme, 'dark');
		assert.equal(localStorage.getItem('visio-docs-theme'), 'dark');
		assert.equal(button.getAttribute('aria-label'), 'Switch to light theme');
		button.click();
		assert.equal(document.documentElement.dataset.theme, 'light');
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
});
