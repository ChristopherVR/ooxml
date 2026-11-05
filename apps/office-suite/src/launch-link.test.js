import assert from 'node:assert/strict';
import { test } from 'node:test';
import { isOneDriveLink, resolveLaunchLinks, takeLaunchLink } from './launch-link.js';

const base = 'https://example.github.io/ooxml/';
const share = 'https://1drv.ms/p/c/abc/xyz';

test('accepts OneDrive and SharePoint https links only', () => {
	assert.equal(isOneDriveLink(share), true);
	assert.equal(isOneDriveLink('https://contoso.sharepoint.com/:p:/r/x'), true);
	assert.equal(isOneDriveLink('http://1drv.ms/p/x'), false);
	assert.equal(isOneDriveLink('https://1drv.ms.evil.example/x'), false);
	assert.equal(isOneDriveLink('not a link'), false);
});

test('reads ?open=, strips it from the address and keeps the route', () => {
	const policy = resolveLaunchLinks(undefined, 'onedrive');
	const taken = takeLaunchLink(policy, `${base}?open=${encodeURIComponent(share)}&suite#/files`);
	assert.deepEqual(taken, { link: share, rejected: false, href: `${base}?suite=#/files` });
	assert.equal(takeLaunchLink(policy, base), null);
});

test('reads the web+ooxml protocol and shared text', () => {
	const policy = resolveLaunchLinks(true, 'onedrive');
	const protocol = takeLaunchLink(
		policy,
		`${base}?open=${encodeURIComponent(`web+ooxml:${share}`)}`,
	);
	assert.equal(protocol.link, share);
	const text = takeLaunchLink(
		policy,
		`${base}?title=Deck&text=${encodeURIComponent(`See ${share}`)}`,
	);
	assert.deepEqual(text, { link: share, rejected: false, href: base });
});

test('rejects links the policy does not accept', () => {
	const policy = resolveLaunchLinks(undefined, 'onedrive');
	const taken = takeLaunchLink(
		policy,
		`${base}?open=${encodeURIComponent('https://evil.example/x')}`,
	);
	assert.deepEqual(taken, { link: null, rejected: true, href: base });
});

test('is configurable: off, renamed, custom acceptance; custom providers opt in', () => {
	assert.equal(resolveLaunchLinks(false, 'onedrive'), null);
	assert.equal(resolveLaunchLinks(undefined, 'dropbox'), null);
	const policy = resolveLaunchLinks(
		{ param: 'file', accept: (link) => link.startsWith('box:') },
		'box',
	);
	assert.equal(takeLaunchLink(policy, `${base}?file=box:42`).link, 'box:42');
	assert.equal(takeLaunchLink(policy, `${base}?open=box:42`), null);
});
