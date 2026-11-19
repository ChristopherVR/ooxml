import assert from 'node:assert/strict';
import { test } from 'node:test';

import { lockInput } from './session-lock.js';

/** The part of an element the lock touches. */
class FakeHost extends EventTarget {
	attributes = new Map();
	setAttribute(name, value) {
		this.attributes.set(name, value);
	}
	removeAttribute(name) {
		this.attributes.delete(name);
	}
}

function press(target, key) {
	const event = new Event('keydown', { cancelable: true });
	event.key = key;
	target.dispatchEvent(event);
	return event;
}

test('a lock refuses input and marks the editor busy until released', () => {
	const host = new FakeHost();
	const unlock = lockInput(host);
	assert.equal(host.attributes.get('aria-busy'), 'true');
	assert.equal(press(host, 'x').defaultPrevented, true);
	unlock();
	assert.equal(host.attributes.has('aria-busy'), false);
	assert.equal(press(host, 'x').defaultPrevented, false);
});

test('Tab still moves focus while locked', () => {
	const host = new FakeHost();
	lockInput(host);
	assert.equal(press(host, 'Tab').defaultPrevented, false);
});

test('overlapping locks hold until the last is released, and releasing twice is harmless', () => {
	const host = new FakeHost();
	const first = lockInput(host);
	const second = lockInput(host);
	first();
	first();
	assert.equal(press(host, 'x').defaultPrevented, true);
	second();
	assert.equal(press(host, 'x').defaultPrevented, false);
});
