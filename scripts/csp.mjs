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
const ANY_META = /<meta\s+http-equiv=["']Content-Security-Policy["'][^>]*>\s*/giu;
const CHARSET = /<meta\s+charset=["']?[\w-]+["']?\s*\/?>/iu;
const SCRIPT = /<script\b([^>]*)>([\s\S]*?)<\/script\s*>/giu;
// Script elements that hold data, not code: the policy does not apply to them.
const DATA_TYPE = /\btype\s*=\s*["']?(application\/(ld\+)?json|text\/template)\b/iu;
const HAS_SRC = /\bsrc\s*=/iu;
const TAG = /<[a-z][\w-]*\b[^>]*>/giu;
const HANDLER = /\son[a-z]+\s*=/iu;
const SCRIPT_URL = /\s(?:href|src|action|formaction)\s*=\s*["']?\s*javascript:/iu;
const FOLDER = 'csp-inline';

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
	if (!contentSecurityPolicyProblems(html).length) return html;
	const scripted = html.replace(SCRIPT, (match, attributes, body) => {
		if (HAS_SRC.test(attributes) || DATA_TYPE.test(attributes) || !body.trim()) return match;
		return `<script${attributes} src="${externalize(body)}"></script>`;
	});
	const bare = scripted.replace(ANY_META, '');
	const charset = CHARSET.exec(bare);
	if (charset) {
		const end = charset.index + charset[0].length;
		return `${bare.slice(0, end)}\n${META}${bare.slice(end)}`;
	}
	return bare.replace(/<head\b[^>]*>/iu, (head) => `${head}\n${META}`);
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
	const problems = [];
	const metas = html.match(ANY_META) ?? [];
	if (metas.length !== 1) problems.push(`has ${metas.length} policies, expected 1`);
	else if (metas[0].trim() !== META) problems.push('has a different policy');
	const head = html.slice(0, html.search(/<\/head>/iu) + 1 || undefined);
	const first = head.search(/<(script|link|style|base)\b/iu);
	if (metas.length && first !== -1 && first < head.indexOf(metas[0].trim()))
		problems.push('loads something before the policy');
	for (const [, attributes, body] of html.matchAll(SCRIPT))
		if (!HAS_SRC.test(attributes) && !DATA_TYPE.test(attributes) && body.trim())
			problems.push(`has an inline script: ${body.trim().slice(0, 60)}`);
	const withoutScripts = html.replace(SCRIPT, '');
	for (const [tag] of withoutScripts.matchAll(TAG)) {
		if (HANDLER.test(tag)) problems.push(`has an inline event handler: ${tag.slice(0, 80)}`);
		if (SCRIPT_URL.test(tag)) problems.push(`has a javascript: URL: ${tag.slice(0, 80)}`);
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
