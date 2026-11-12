import assert from 'node:assert/strict';
import { test } from 'node:test';

import { positionLabel, pressedState, splitFileName } from './document-bar.js';

test('a file name splits at its last dot, so the bar can mute the extension', () => {
	assert.deepEqual(splitFileName('Quarterly review.pptx'), ['Quarterly review', '.pptx']);
	assert.deepEqual(splitFileName('v1.2 notes.docx'), ['v1.2 notes', '.docx']);
	assert.deepEqual(splitFileName('README'), ['README', '']);
	assert.deepEqual(splitFileName('.hidden'), ['.hidden', '']);
});

test('the position reads as one-based and never runs past the end', () => {
	assert.equal(positionLabel({ unit: 'Slide', index: 0, count: 9 }), 'Slide 1 / 9');
	assert.equal(positionLabel({ unit: 'Page', index: 4, count: 2 }), 'Page 2 / 2');
	assert.equal(positionLabel({ unit: 'Slide', index: 0, count: 0 }), '');
});

test('a formatting button is pressed, not pressed, or mixed for a selection that has both', () => {
	assert.equal(pressedState(true), 'true');
	assert.equal(pressedState(false), 'false');
	assert.equal(pressedState(undefined), 'false');
	assert.equal(pressedState(null), 'mixed');
});
