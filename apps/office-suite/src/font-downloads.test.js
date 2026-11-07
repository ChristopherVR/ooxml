import assert from 'node:assert/strict';
import { test } from 'node:test';

import {
	fontDownloads,
	editorGoogleFonts,
	googleFontsOption,
	normalizeFontServer,
	setFontDownloads,
} from './font-downloads.js';

test('downloads default to Google Fonts, which keeps the editors default', () => {
	assert.deepEqual(fontDownloads(), { mode: 'google', server: '' });
	assert.equal(googleFontsOption(), undefined);
});

test('off and a custom server map onto the editors googleFonts option', () => {
	assert.equal(googleFontsOption({ mode: 'off', server: '' }), false);
	assert.deepEqual(
		googleFontsOption({ mode: 'custom', server: 'https://fonts.example.com/css2' }),
		{
			baseUrl: 'https://fonts.example.com/css2',
		},
	);
});

test('only absolute http(s) addresses without credentials, query or fragment are accepted', () => {
	assert.equal(
		normalizeFontServer(' https://fonts.example.com/css2 '),
		'https://fonts.example.com/css2',
	);
	for (const bad of [
		'',
		'fonts.example.com',
		'ftp://fonts.example.com/css2',
		'javascript:alert(1)',
		'https://user:pw@fonts.example.com/css2',
		'https://fonts.example.com/css2?family=x',
		'https://fonts.example.com/css2#x',
		42,
	]) {
		assert.equal(normalizeFontServer(bad), '', String(bad));
	}
});

test('a custom choice without a valid server turns downloads off instead of using Google', () => {
	assert.deepEqual(setFontDownloads({ mode: 'custom', server: 'not a url' }), {
		mode: 'off',
		server: '',
	});
	assert.deepEqual(setFontDownloads({ mode: 'custom', server: 'https://fonts.example.com/css2' }), {
		mode: 'custom',
		server: 'https://fonts.example.com/css2',
	});
	assert.deepEqual(setFontDownloads({ mode: 'google' }), {
		mode: 'google',
		server: 'https://fonts.example.com/css2',
	});
});

test('Word, Excel and Visio get the Google choice explicitly, since unset means off for them', () => {
	assert.deepEqual(editorGoogleFonts({ mode: 'google', server: '' }), {
		baseUrl: 'https://fonts.googleapis.com/css2',
	});
	assert.equal(editorGoogleFonts({ mode: 'off', server: '' }), false);
	assert.deepEqual(editorGoogleFonts({ mode: 'custom', server: 'https://f.example/css2' }), {
		baseUrl: 'https://f.example/css2',
	});
});
