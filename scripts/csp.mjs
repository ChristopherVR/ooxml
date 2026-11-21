/**
 * csp.mjs: the one Content Security Policy of every page this repository publishes (the suite,
 * its 404 page, every viewer's documentation and demos on GitHub Pages) and of the desktop app.
 *
 * The Pages origin, christophervr.github.io, is shared with every other project site on it, so
 * every page carries the same script policy: script only from this origin, never inline. GitHub
 * Pages cannot send headers, so the policy is a meta tag that `applyContentSecurityPolicy` writes
 * into each built page, after moving any inline script a page builder emitted into its own file.
 * `checkContentSecurityPolicy` fails the build on a page without it. The desktop app sets the
 * same string in `desktop/src-tauri/tauri.conf.json`; `scripts/csp.test.mjs` keeps the two equal.
 *
 * Why each source is here (PRIVACY.md lists the requests):
 * - connect-src stays broad: the assistant, team and collaboration servers are user-configured
 *   endpoints. `ipc:` and `http://ipc.localhost` are the desktop shell's IPC channel. `blob:` and
 *   `data:` are documents and images the editors read back (embedded fonts, canvas exports).
 * - font-src: `blob:` serves fonts embedded in a document; fonts.gstatic.com serves fonts a
 *   document uses but neither the device nor the file has (docs/font-loading.md).
 * - frame-src: this site's own pages, and the players of the online-video hosts a deck may link
 *   (YouTube and Vimeo). Arbitrary websites are not framed, so a Teams website tab offers
 *   "Open externally" instead of showing the site inside the app.
 * - style-src keeps 'unsafe-inline': the editors position and colour content with inline styles.
 */
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { html as htmlSyntax, parse } from 'parse5';
import { escapeHtml } from './release-markdown.mjs';

export const POLICY = {
	'default-src': ["'self'"],
	'script-src': ["'self'"],
	'object-src': ["'none'"],
	'base-uri': ["'self'"],
	'form-action': ["'self'"],
	'img-src': ["'self'", 'data:', 'blob:', 'https:'],
	'media-src': ["'self'", 'data:', 'blob:', 'https:'],
	'font-src': ["'self'", 'data:', 'blob:', 'https://fonts.gstatic.com'],
	'style-src': ["'self'", "'unsafe-inline'", 'https://fonts.googleapis.com'],
	'connect-src': [
		"'self'",
		'blob:',
		'data:',
		'ipc:',
		'http://ipc.localhost',
		'https:',
		'wss:',
		'ws://localhost:*',
		'ws://127.0.0.1:*',
		'http://localhost:*',
		'http://127.0.0.1:*',
	],
	'worker-src': ["'self'", 'blob:'],
	'frame-src': [
		"'self'",
		'https://christophervr.github.io',
		'https://www.youtube.com',
		'https://player.vimeo.com',
	],
	'manifest-src': ["'self'"],
};

export const CONTENT_SECURITY_POLICY = Object.entries(POLICY)
	.map(([directive, sources]) => `${directive} ${sources.join(' ')}`)
	.join('; ');

const META = `<meta http-equiv="Content-Security-Policy" content="${CONTENT_SECURITY_POLICY}" />`;
const FOLDER = 'csp-inline';
const DATA_TYPES = new Set(['application/json', 'application/ld+json', 'text/template']);
const EARLY_LOADS = new Set(['script', 'link', 'style', 'base']);
const URL_ATTRIBUTES = new Set(['href', 'src', 'action', 'formaction']);

function* elements(node) {
	if (node.tagName) yield node;
	for (const child of node.childNodes ?? []) yield* elements(child);
}

const attribute = (node, name) => node.attrs.find((attr) => attr.name === name)?.value;
const isHtml = (node) => node.namespaceURI === htmlSyntax.NS.HTML;
const isPolicy = (node) =>
	isHtml(node) &&
	node.tagName === 'meta' &&
	attribute(node, 'http-equiv')?.toLowerCase() === 'content-security-policy';

function inlineScript(node) {
	if (node.tagName !== 'script') return '';
	if (isHtml(node)) {
		const type = attribute(node, 'type')?.split(';', 1)[0].trim().toLowerCase();
		if (attribute(node, 'src') !== undefined || DATA_TYPES.has(type)) return '';
	} else if (attribute(node, 'href') !== undefined) return '';
	return node.childNodes.map((child) => child.value ?? '').join('');
}

function forbiddenUrlScheme(node, value) {
	const scheme = URL.parse(value)?.protocol;
	switch (scheme) {
		case 'javascript:':
		case 'vbscript:':
			return scheme;
		case 'data:':
			// The policy allows data URLs for images, media and fonts, but forbids them for scripts.
			return node.tagName === 'script' ? scheme : undefined;
	}
}

function inspect(html) {
	const nodes = [...elements(parse(html, { sourceCodeLocationInfo: true }))];
	const head = nodes.find((node) => node.tagName === 'head' && isHtml(node));
	const policies = nodes.filter(isPolicy);
	const firstLoadAt = nodes.reduce(
		(first, node) =>
			EARLY_LOADS.has(node.tagName)
				? Math.min(first, node.sourceCodeLocation?.startOffset ?? Infinity)
				: first,
		Infinity,
	);
	return { nodes, head, policies, firstLoadAt };
}

function htmlFiles(dir, out = []) {
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		const path = join(dir, entry.name);
		if (entry.isDirectory()) htmlFiles(path, out);
		else if (entry.name.endsWith('.html')) out.push(path);
	}
	return out;
}

/**
 * Write the policy into `html`, first after its charset declaration (a meta policy covers only
 * what follows it), replacing any policy the page already had. Inline code is handed to
 * `externalize`, which returns the URL of a file holding it. A page that already passes is
 * returned unchanged, so applying the policy twice changes nothing.
 */
export function withContentSecurityPolicy(html, externalize) {
	const document = inspect(html);
	if (!policyProblems(html, document).length) return html;
	const { nodes, head, policies, firstLoadAt } = document;
	const charset = head?.childNodes.find(
		(node) =>
			node.tagName === 'meta' &&
			!isPolicy(node) &&
			attribute(node, 'charset') !== undefined &&
			node.sourceCodeLocation.endOffset <= firstLoadAt,
	);
	const insertAt =
		charset?.sourceCodeLocation.endOffset ?? head?.sourceCodeLocation?.startTag?.endOffset;
	if (insertAt === undefined) return html;
	const edits = [{ start: insertAt, end: insertAt, text: `\n${META}` }];
	for (const { sourceCodeLocation: loc } of policies) {
		edits.push({ start: loc.startOffset, end: loc.endOffset, text: '' });
	}
	for (const node of nodes) {
		if (!isHtml(node) || !inlineScript(node).trim()) continue;
		const { startTag, endTag, endOffset } = node.sourceCodeLocation;
		const body = html.slice(startTag.endOffset, endTag?.startOffset ?? endOffset);
		const start = startTag.startOffset + '<script'.length;
		edits.push({ start, end: start, text: ` src="${escapeHtml(externalize(body))}"` });
		edits.push({ start: startTag.endOffset, end: endTag?.startOffset ?? endOffset, text: '' });
	}
	edits.sort((a, b) => a.start - b.start || a.end - b.end);
	let result = '';
	let offset = 0;
	for (const { start, end, text } of edits) {
		result += html.slice(offset, start) + text;
		offset = end;
	}
	return result + html.slice(offset);
}

/**
 * Apply the policy to every page under `dir`. Inline scripts move to `<dir>/csp-inline/`, named by
 * their content hash and referenced from the site root `base` (the URL path `dir` is served at),
 * so a page that sets `<base href>` still finds them.
 */
export function applyContentSecurityPolicy(dir, { base = '/' } = {}) {
	const root = base.endsWith('/') ? base : `${base}/`;
	for (const file of htmlFiles(dir)) {
		const html = readFileSync(file, 'utf8');
		const next = withContentSecurityPolicy(html, (body) => {
			const name = `${createHash('sha256').update(body).digest('hex').slice(0, 16)}.js`;
			mkdirSync(join(dir, FOLDER), { recursive: true });
			writeFileSync(join(dir, FOLDER, name), body);
			return `${root}${FOLDER}/${name}`;
		});
		if (next !== html) writeFileSync(file, next);
	}
}

/** What is wrong with one page, against the policy. */
export function contentSecurityPolicyProblems(html) {
	return policyProblems(html, inspect(html));
}

function policyProblems(html, { nodes, head, policies, firstLoadAt }) {
	const problems = [];
	if (policies.length !== 1) problems.push(`has ${policies.length} policies, expected 1`);
	else if (attribute(policies[0], 'content') !== CONTENT_SECURITY_POLICY)
		problems.push('has a different policy');
	else if (policies[0].parentNode !== head) problems.push('has a policy outside the head');
	if (policies.length && firstLoadAt < policies[0].sourceCodeLocation.startOffset)
		problems.push('loads something before the policy');
	for (const node of nodes) {
		const body = inlineScript(node).trim();
		if (body) problems.push(`has an inline script: ${body.slice(0, 60)}`);
		const start = node.sourceCodeLocation?.startOffset;
		const tag = start === undefined ? node.tagName : html.slice(start, start + 80);
		if (node.attrs.some(({ name }) => /^on[a-z]+$/u.test(name)))
			problems.push(`has an inline event handler: ${tag}`);
		for (const { name, value } of node.attrs) {
			const scheme = URL_ATTRIBUTES.has(name) ? forbiddenUrlScheme(node, value) : undefined;
			if (scheme) problems.push(`has a ${scheme} URL: ${tag}`);
		}
	}
	return problems;
}

/** Every page under `dir` that breaks the policy, with what is wrong; empty when all pass. */
export function checkContentSecurityPolicy(dir) {
	const failures = [];
	for (const file of htmlFiles(dir)) {
		const problems = contentSecurityPolicyProblems(readFileSync(file, 'utf8'));
		if (problems.length) failures.push({ page: relative(dir, file), problems });
	}
	return failures;
}

/** Throws, naming each page, when a page under `dir` lacks the policy or would break under it. */
export function assertContentSecurityPolicy(dir) {
	const failures = checkContentSecurityPolicy(dir);
	if (!failures.length) return;
	const lines = failures
		.slice(0, 40)
		.map(({ page, problems }) => `  ${page}: ${problems.slice(0, 3).join('; ')}`);
	throw new Error(
		`${failures.length} published page(s) break the Content Security Policy:\n${lines.join('\n')}`,
	);
}

// node scripts/csp.mjs <dir>: check a built site (CI runs it over pages-dist).
if (import.meta.main) {
	const dir = process.argv[2];
	if (!dir) {
		console.log(CONTENT_SECURITY_POLICY);
	} else {
		assertContentSecurityPolicy(dir);
		console.log(`Every page under ${dir} carries the Content Security Policy.`);
	}
}
