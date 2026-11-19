import assert from 'node:assert/strict';
import { test } from 'node:test';

import { isLocked, lockInput, takeComposed, takeQueued } from './session-lock.js';

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

function press(target, key, modifiers = {}) {
	const event = new Event('keydown', { cancelable: true });
	Object.assign(event, { key, ...modifiers });
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

function compose(target, data) {
	const event = new Event('compositionend');
	event.data = data;
	target.dispatchEvent(event);
}

test('text an input method commits while locked is recorded once, and only while locked', () => {
	const host = new FakeHost();
	compose(host, 'a');
	const unlock = lockInput(host);
	assert.equal(isLocked(host), true);
	compose(host, 'あ');
	compose(host, 'い');
	assert.equal(unlock(), true);
	assert.equal(isLocked(host), false);
	compose(host, 'b');
	assert.equal(takeComposed(host), 'あい');
	assert.equal(takeComposed(host), '');
});

test('only the last release of overlapping locks reports that the editor is free', () => {
	const host = new FakeHost();
	const first = lockInput(host);
	const second = lockInput(host);
	assert.equal(first(), false);
	assert.equal(second(), true);
	assert.equal(second(), false);
});

test('Ctrl+S or Cmd+S pressed while locked is refused and queued once, other keys are not', () => {
	const host = new FakeHost();
	const unlock = lockInput(host);
	assert.equal(press(host, 's', { ctrlKey: true }).defaultPrevented, true);
	press(host, 's', { metaKey: true });
	press(host, 's');
	press(host, 'z', { ctrlKey: true });
	unlock();
	assert.deepEqual(takeQueued(host), ['save']);
	assert.deepEqual(takeQueued(host), []);
	press(host, 's', { ctrlKey: true });
	assert.deepEqual(takeQueued(host), [], 'an unlocked editor saves the usual way');
});
