import assert from 'node:assert/strict';
import { test } from 'node:test';

import { avatar } from './profile.js';

test('an avatar shows initials without text, so its button is named by the profile name', () => {
	const html = avatar({ name: 'Alex Morgan', photo: '' });
	assert.equal(html, '<span class="avatar" aria-hidden="true" data-initials="AM"></span>');
});

test('an avatar with a photo shows the photo and no initials', () => {
	const html = avatar({ name: 'Alex Morgan', photo: 'data:image/png;base64,AAAA' }, true);
	assert.doesNotMatch(html, /data-initials/u);
	assert.match(html, /<img src="data:image\/png;base64,AAAA" alt="" \/>/u);
});
