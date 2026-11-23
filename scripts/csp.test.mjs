import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import {
	CONTENT_SECURITY_POLICY,
	applyContentSecurityPolicy,
	checkContentSecurityPolicy,
	contentSecurityPolicyProblems,
	withContentSecurityPolicy,
} from './csp.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const META = `<meta http-equiv="Content-Security-Policy" content="${CONTENT_SECURITY_POLICY}" />`;
const page = (head, body = '') =>
	`<!doctype html><html><head><meta charset="utf-8" />${head}</head><body>${body}</body></html>`;

test('the desktop app carries the same policy as the published pages', () => {
	const config = JSON.parse(readFileSync(join(ROOT, 'desktop/src-tauri/tauri.conf.json'), 'utf8'));
	assert.equal(config.app.security.csp, CONTENT_SECURITY_POLICY);
});

test('the policy keeps script to this origin and allows what the editors load', () => {
	const directives = Object.fromEntries(
		CONTENT_SECURITY_POLICY.split('; ').map((d) => [d.split(' ')[0], d.split(' ').slice(1)]),
	);
	assert.deepEqual(directives['script-src'], ["'self'"]);
	assert.deepEqual(directives['object-src'], ["'none'"]);
	assert.ok(directives['font-src'].includes('blob:'), 'fonts embedded in a document');
	assert.ok(directives['connect-src'].includes('blob:'), 'reading back blob URLs');
	assert.deepEqual(directives['frame-src'], [
		"'self'",
		'https://christophervr.github.io',
		'https://www.youtube.com',
		'https://player.vimeo.com',
		'https://forms.office.com',
		'https://forms.cloud.microsoft',
	]);
});

test('the site sources leave the policy to the build', () => {
	for (const name of ['index.html', '404.html']) {
		const html = readFileSync(join(ROOT, 'site', name), 'utf8');
		assert.doesNotMatch(html, /http-equiv="Content-Security-Policy"/u, name);
	}
});

test('the policy is written first in the head, replacing any other', () => {
	const html = withContentSecurityPolicy(
		page(
			'<meta http-equiv="Content-Security-Policy" content="default-src *" /><link rel="stylesheet" href="a.css" />',
		),
		() => assert.fail('no inline script'),
	);
	assert.equal(html.match(/Content-Security-Policy/gu).length, 1);
	assert.ok(html.indexOf(META) > html.indexOf('charset'));
	assert.ok(html.indexOf(META) < html.indexOf('<link'));
	assert.deepEqual(contentSecurityPolicyProblems(html), []);
});

test('inline scripts move to files, in order, keeping their attributes', () => {
	const moved = [];
	const html = withContentSecurityPolicy(
		page(
			'<script id="theme">document.documentElement.className = "dark"</script><script type="application/json" id="data">{"a":1}</script>',
			'<script type="module">import "./x.js"</script>',
		),
		(body) => {
			moved.push(body);
			return `/s/${moved.length}.js`;
		},
	);
	assert.deepEqual(moved, ['document.documentElement.className = "dark"', 'import "./x.js"']);
	assert.match(html, /<script src="\/s\/1\.js" id="theme"><\/script>/u);
	assert.match(html, /<script src="\/s\/2\.js" type="module"><\/script>/u);
	assert.match(html, /<script type="application\/json" id="data">\{"a":1\}<\/script>/u);
	assert.deepEqual(contentSecurityPolicyProblems(html), []);
	assert.equal(
		withContentSecurityPolicy(html, () => assert.fail('already applied')),
		html,
	);
});

test('recognizes actual script and policy elements and preserves other markup', () => {
	const inert = '<!-- <script>example()</script> -->';
	const source = page(
		`${inert}<META content="default-src *" HTTP-EQUIV=Content-Security-Policy>` +
			'<SCRIPT data-label="a > b" data-src="example">example()</SCRIPT >',
	);
	const next = withContentSecurityPolicy(source, (body) => {
		assert.equal(body, 'example()');
		return '/inline.js?x=1&y=2';
	});
	assert.ok(next.includes(inert));
	assert.ok(
		next.includes(
			'<SCRIPT src="/inline.js?x=1&amp;y=2" data-label="a > b" data-src="example"></SCRIPT >',
		),
	);
	assert.deepEqual(contentSecurityPolicyProblems(next), []);
});

test('the check names pages the policy would break', () => {
	assert.deepEqual(contentSecurityPolicyProblems(page('')), ['has 0 policies, expected 1']);
	const broken = page(META, '<button onclick="go()">Go</button><a href="javascript:go()">Go</a>');
	assert.deepEqual(
		contentSecurityPolicyProblems(broken).map((problem) => problem.split(':')[0]),
		['has an inline event handler', 'has a javascript'],
	);
	assert.deepEqual(contentSecurityPolicyProblems(page(META, '<script>go()</script>')), [
		'has an inline script: go()',
	]);
	assert.deepEqual(contentSecurityPolicyProblems(page('<script src="a.js"></script>' + META)), [
		'loads something before the policy',
	]);
});

test('the check allows data images and reports executable URL schemes', () => {
	const html = page(
		META,
		'<img src="data:image/png;base64,AAAA"><script src="data:text/javascript,0"></script><a href="vbscript:example">Example</a>',
	);
	assert.deepEqual(
		contentSecurityPolicyProblems(html).map((problem) => problem.split(' URL:')[0]),
		['has a data:', 'has a vbscript:'],
	);
});

test('applying the policy to a built site covers every page and passes the check', () => {
	const dir = mkdtempSync(join(tmpdir(), 'csp-'));
	try {
		mkdirSync(join(dir, 'docs/guide'), { recursive: true });
		writeFileSync(join(dir, 'index.html'), page('<script src="suite.js"></script>'));
		writeFileSync(join(dir, 'docs/guide/a.html'), page('<script>window.__DATA__ = 1</script>'));
		assert.equal(checkContentSecurityPolicy(dir).length, 2);
		applyContentSecurityPolicy(dir, { base: '/ooxml/' });
		assert.deepEqual(checkContentSecurityPolicy(dir), []);
		const [file] = readdirSync(join(dir, 'csp-inline'));
		assert.equal(readFileSync(join(dir, 'csp-inline', file), 'utf8'), 'window.__DATA__ = 1');
		assert.match(
			readFileSync(join(dir, 'docs/guide/a.html'), 'utf8'),
			new RegExp(`<script src="/ooxml/csp-inline/${file}"></script>`, 'u'),
		);
	} finally {
		rmSync(dir, { recursive: true, force: true });
	}
});
