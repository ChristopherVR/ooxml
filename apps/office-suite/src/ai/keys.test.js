import assert from 'node:assert/strict';
import { test } from 'node:test';

import { installStorage } from './test-storage.js';

installStorage();
const { forgetRememberedKeys } = await import('./keys.js');
const { askForKey } = await import('./settings.js');
const { currentConnection, saveUserConnection } = await import('./providers.js');
const { loadPolicy } = await import('./policy.js');
// A configuration turns the assistant on, so users may add their own models.
globalThis.ooxmlAiConfig = {};
await loadPolicy();
delete globalThis.ooxmlAiConfig;

const connection = {
	id: 'own-1',
	label: 'My model',
	provider: 'anthropic',
	endpoint: 'https://api.anthropic.com/v1/messages',
};

/** Runs `steps` with just enough DOM for one dialog: the form it builds and the dialog it opens. */
function withDialog(steps) {
	const seen = { form: null, closed: false };
	const nodes = {
		dialog: { showModal() {}, close: () => (seen.closed = true) },
		'dialog-title': {},
		'dialog-content': { replaceChildren: (form) => (seen.form = form) },
	};
	const { FormData } = globalThis;
	globalThis.document = { createElement: () => ({}), getElementById: (id) => nodes[id] };
	globalThis.FormData = class {
		constructor(form) {
			this.values = form.values ?? {};
		}
		get(name) {
			return this.values[name] ?? null;
		}
		has(name) {
			return name in this.values;
		}
	};
	try {
		steps(seen);
	} finally {
		delete globalThis.document;
		globalThis.FormData = FormData;
	}
	return seen;
}

test('the key dialog offers no remember option and says the key is kept for this tab', () => {
	const { form } = withDialog(() => askForKey(connection));
	assert.doesNotMatch(form.innerHTML, /remember/iu);
	assert.match(form.innerHTML, /kept only for this tab and is cleared when you close it/u);
});

test('a key entered in the dialog lands in session storage only', () => {
	const seen = withDialog((dialog) => {
		askForKey(connection);
		dialog.form.values = { key: 'sk-dialog-secret' };
		dialog.form.onsubmit({ preventDefault() {} });
	});
	assert.equal(seen.closed, true);
	assert.match(sessionStorage.dump(), /sk-dialog-secret/u);
	assert.doesNotMatch(localStorage.dump(), /sk-dialog-secret/u);
});

test('a key saved with a new model is used but never written to local storage', () => {
	saveUserConnection({ ...connection, id: 'own-2', model: 'm' }, 'sk-own-secret');
	assert.equal(currentConnection().key, 'sk-own-secret');
	assert.doesNotMatch(localStorage.dump(), /sk-own-secret/u);
});

test('loading forgets every key an earlier version remembered in local storage', () => {
	localStorage.setItem('ooxml-suite-ai-keys-v2-default', JSON.stringify({ 'own-2': 'sk-old-1' }));
	localStorage.setItem('ooxml-suite-ai-keys-v2-work', JSON.stringify({ x: 'sk-old-2' }));
	localStorage.setItem(
		'ooxml-suite-ai-v1-default',
		JSON.stringify({ provider: 'openai', endpoint: 'https://e', key: 'sk-old-3', remember: true }),
	);
	localStorage.setItem('ooxml-suite-theme', 'dark');
	forgetRememberedKeys();
	assert.doesNotMatch(localStorage.dump(), /sk-old|remember/u);
	assert.equal(localStorage.getItem('ooxml-suite-ai-keys-v2-work'), null);
	assert.deepEqual(JSON.parse(localStorage.getItem('ooxml-suite-ai-v1-default')), {
		provider: 'openai',
		endpoint: 'https://e',
	});
	assert.equal(localStorage.getItem('ooxml-suite-theme'), 'dark');
	// A key in this tab's session storage is not a remembered one.
	assert.equal(currentConnection().key, 'sk-own-secret');
});
