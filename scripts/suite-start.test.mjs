import assert from 'node:assert/strict';
import { test } from 'node:test';

test('the chooser opens bundled preview apps while the full site opens product sites', async (t) => {
	const keys = ['document', 'location', 'localStorage', 'matchMedia', 'window'];
	const originals = keys.map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)]);
	t.after(() => {
		for (const [key, descriptor] of originals) {
			if (descriptor) Object.defineProperty(globalThis, key, descriptor);
			else delete globalThis[key];
		}
	});
	let preview = true;
	let click;
	let markup;
	let assigned;
	const themeButton = { setAttribute() {}, getAttribute() {} };
	globalThis.document = {
		body: { dataset: {}, append() {} },
		getElementById: () => null,
		querySelector: (selector) =>
			selector === 'meta[name="ooxml-preview"]' ? (preview ? {} : null) : { content: './' },
		addEventListener() {},
		removeEventListener() {},
		createElement: () => ({
			setAttribute() {},
			set innerHTML(value) {
				markup = value;
			},
			querySelector: (selector) =>
				selector === '.start-theme' ? themeButton : { addEventListener() {}, focus() {} },
			addEventListener: (type, handler) => {
				if (type === 'click') click = handler;
			},
		}),
	};
	globalThis.location = {
		href: 'https://preview.example/ooxml/',
		assign: (url) => {
			assigned = url.href;
		},
	};
	globalThis.localStorage = { getItem: () => null };
	globalThis.window = { addEventListener() {} };
	globalThis.matchMedia = () => ({ matches: false, addEventListener() {} });
	const { showChooser } = await import('../apps/office-suite/src/start.js');
	for (const mode of [true, false]) {
		preview = mode;
		showChooser();
		assert.match(markup, mode ? /open one app separately/ : /guides and demos/);
		for (const [id, site, app] of [
			['docx', 'docx', 'word'],
			['xlsx', 'xlsx', 'excel'],
			['pptx', 'pptx', 'powerpoint'],
			['vsdx', 'visio', 'visio'],
			['teams', 'teams', 'teams'],
			['docs', 'docs', null],
		]) {
			click({ target: { closest: () => ({ dataset: { start: id } }) } });
			const path = mode && app ? `apps/${app}` : site;
			assert.equal(assigned, `https://preview.example/ooxml/${path}/`);
		}
	}
});
