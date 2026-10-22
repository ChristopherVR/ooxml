/**
 * Weekly documentation health check: finds the mechanically detectable kinds of drift in the
 * repository's Markdown, with no network and no dependencies.
 *
 *   a. dead relative links and image references (and, as warnings, dead #anchors)
 *   b. repository paths written in inline code that do not exist
 *   c. stale terms from scripts/check-docs.config.json
 *   d. VitePress sidebar and nav links that point at no page (errors), and English pages that no
 *      sidebar reaches (warnings)
 *   e. translation lag per viewer docs site (warnings)
 *
 * Usage: node scripts/check-docs.mjs [--out report.md] [--warn-only] [--root dir]
 * Exit code 1 when there are errors, unless --warn-only.
 */
import { execFileSync } from 'node:child_process';
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, posix, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const SKIP_FILE = /(?:^|\/)(?:CHANGELOG\.md|node_modules|__fixtures__|__snapshots__)(?:\/|$)/;
const REPO_DIRS = 'src|docs|scripts|viewers|demos|e2e|apps|mcp|site|\\.github';
const LOCALE = /^[a-z]{2}(?:-[A-Za-z]{2,4})?$/;
const VIEWERS = ['pptx', 'docx', 'xlsx', 'visio', 'teams'];

export function globToRegExp(glob) {
	let out = '';
	for (let i = 0; i < glob.length; i++) {
		const c = glob[i];
		if (c === '*' && glob[i + 1] === '*') {
			out += '.*';
			i++;
			if (glob[i + 1] === '/') i++;
		} else if (c === '*') out += '[^/]*';
		else if (c === '?') out += '[^/]';
		else out += c.replace(/[.+^${}()|[\]\\]/g, '\\$&');
	}
	return new RegExp(`^${out}$`);
}

/** Lines of a Markdown file with fenced code flagged and inline code blanked out. */
export function scanLines(text) {
	const rows = [];
	let fence = null;
	text.split(/\r?\n/).forEach((raw, i) => {
		const m = /^\s*(`{3,}|~{3,})/.exec(raw);
		if (fence) {
			if (m && m[1][0] === fence[0] && m[1].length >= fence.length) fence = null;
			rows.push({ n: i + 1, raw, prose: '', fenced: true });
			return;
		}
		if (m) {
			fence = m[1];
			rows.push({ n: i + 1, raw, prose: '', fenced: true });
			return;
		}
		const prose = raw.replace(/(`+)(.+?)\1/g, (all) => 'x'.repeat(all.length));
		rows.push({ n: i + 1, raw, prose, fenced: false });
	});
	return rows;
}

function slugs(heading) {
	const text = heading
		.replace(/\s*\{#[^}]+\}\s*$/, '')
		.replace(/!?\[([^\]]*)\]\([^)]*\)/g, '$1')
		.replace(/<[^>]+>/g, '')
		.replace(/[`*_~]/g, (c) => (c === '_' ? '_' : ''))
		.trim()
		.toLowerCase();
	const github = text.replace(/[^\p{L}\p{N}\s_-]/gu, '').replace(/\s/g, '-');
	const vitepress = text
		.replace(/[\s~`!@#$%^&*()\-_+=[\]{}|\\;:"'<>,.?/]+/g, '-')
		.replace(/^-+|-+$/g, '');
	// VitePress also folds accents (`über` becomes `uber`).
	const folded = vitepress.normalize('NFD').replace(/\p{M}/gu, '');
	return [github, vitepress, folded];
}

export function anchorsOf(text) {
	const found = new Set();
	const seen = new Map();
	for (const row of scanLines(text)) {
		if (row.fenced) continue;
		const h = /^#{1,6}\s+(.*?)\s*#*\s*$/.exec(row.raw);
		if (h) {
			const explicit = /\{#([^}\s]+)\}\s*$/.exec(h[1]);
			if (explicit) found.add(explicit[1]);
			for (const s of slugs(h[1])) {
				const k = seen.get(s) ?? 0;
				seen.set(s, k + 1);
				found.add(k === 0 ? s : `${s}-${k}`);
			}
		}
		for (const id of row.raw.matchAll(/\b(?:id|name)="([^"]+)"/g)) found.add(id[1]);
	}
	return found;
}

/** `viewers/<p>/docs` for a file inside a viewer docs site, else null. */
export function docRootOf(file) {
	const m = /^(viewers\/[^/]+\/docs)\//.exec(file);
	return m ? m[1] : null;
}

function isDir(p) {
	try {
		return statSync(p).isDirectory();
	} catch {
		return false;
	}
}
function isFile(p) {
	try {
		return statSync(p).isFile();
	} catch {
		return false;
	}
}

/** Resolve a link target (no fragment) to a repo-relative file or dir, or null. */
export function resolveTarget(root, file, target, generated = []) {
	const docRoot = docRootOf(file);
	if (target.startsWith('/') && generated.some((r) => r.test(target))) return 'external';
	let base;
	let path = target;
	const site = /^\/ooxml\/([a-z]+)\/(.*)$/.exec(path);
	if (site && VIEWERS.includes(site[1])) {
		base = `viewers/${site[1]}/docs`;
		path = `/${site[2]}`;
	} else if (path.startsWith('/ooxml/')) return 'external';
	else if (path.startsWith('/')) base = docRoot ?? '';
	const joined = path.startsWith('/')
		? posix.join(base, path.slice(1))
		: posix.join(posix.dirname(file), path);
	// A `.html` link in a docs site names the page its `.md` source builds.
	const roots = [joined.replace(/\/+$/, ''), joined.replace(/\.html$/, '')];
	if (base?.endsWith('/docs')) roots.push(posix.join(base, 'public', path.slice(1)));
	for (const p of roots) {
		for (const s of ['', '.md', '.html', '/index.md', '/README.md']) {
			if (isFile(join(root, p + s))) return p + s;
		}
		if (isDir(join(root, p))) return p;
	}
	return null;
}

export function checkLinks(root, file, text, anchorCache, generated = []) {
	const errors = [];
	const warnings = [];
	const own = () => anchorsOf(text);
	for (const row of scanLines(text)) {
		if (row.fenced) continue;
		const targets = [];
		for (const m of row.prose.matchAll(
			/!?\[(?:[^[\]]|\[[^\]]*\])*\]\(\s*<?([^)\s>]+)>?(?:\s+["'][^)]*["'])?\s*\)/g,
		)) {
			targets.push(row.raw.slice(m.index).match(/\]\(\s*<?([^)\s>]+)/)?.[1] ?? m[1]);
		}
		const ref = /^\s*\[[^\]]+\]:\s*(\S+)/.exec(row.prose);
		if (ref) targets.push(row.raw.match(/^\s*\[[^\]]+\]:\s*<?([^\s>]+)/)?.[1] ?? ref[1]);
		for (const m of row.raw.matchAll(/<(?:img|a|source)\b[^>]*?\b(?:src|href)="([^"]+)"/g)) {
			targets.push(m[1]);
		}
		for (const t of targets) {
			if (/^[a-z][a-z0-9+.-]*:/i.test(t) || t.startsWith('//')) continue;
			if (/[{}$<]/.test(t) || t === '') continue;
			const [pathPart, ...frag] = t.split('#');
			// GitHub gives every rendered README an implicit `readme` anchor; `top` is the page top.
			const fragment = /^(?:readme|top)$/i.test(frag.join('#')) ? '' : frag.join('#');
			const path = decodeURI(pathPart.split('?')[0]);
			if (path === '') {
				if (fragment && !/^L\d/.test(fragment) && !own().has(decodeURI(fragment))) {
					warnings.push({
						file,
						line: row.n,
						check: 'anchor',
						message: `No heading for \`#${fragment}\``,
					});
				}
				continue;
			}
			const hit = resolveTarget(root, file, path, generated);
			if (hit === 'external') continue;
			if (hit === null) {
				errors.push({ file, line: row.n, check: 'link', message: `Dead link \`${t}\`` });
				continue;
			}
			if (fragment && hit.endsWith('.md') && isFile(join(root, hit)) && !/^L\d/.test(fragment)) {
				let set = anchorCache.get(hit);
				if (!set) {
					set = anchorsOf(readFileSync(join(root, hit), 'utf8'));
					anchorCache.set(hit, set);
				}
				if (!set.has(decodeURI(fragment))) {
					warnings.push({ file, line: row.n, check: 'anchor', message: `No heading for \`${t}\`` });
				}
			}
		}
	}
	return { errors, warnings };
}

function globExists(root, rel) {
	const parts = rel.split('/');
	const walk = (dir, i) => {
		if (i === parts.length) return true;
		const part = parts[i];
		if (!/[*?]/.test(part)) {
			const next = join(dir, part);
			return existsSync(next) && walk(next, i + 1);
		}
		if (!isDir(dir)) return false;
		const re = globToRegExp(part);
		return readdirSync(dir).some((name) => re.test(name) && walk(join(dir, name), i + 1));
	};
	return walk(root, 0);
}

const BUILD_OUTPUT =
	/(?:^|\/)(?:dist|node_modules|\.vite|coverage|test-results|\.env[\w.]*)(?:\/|$)/;
const PATH_TOKEN = new RegExp(`^(?:\\./)?((?:${REPO_DIRS}|packages)/[\\w.@/*<>{},+-]+)$`);

/** Inline-code paths that look like repository paths. Resolved against the repo root, the file's
 * folders and every ancestor (a viewer README says `src/` for its own src). */
export function checkCodePaths(root, file, text, ignore = {}) {
	const errors = [];
	if ((ignore.files ?? []).some((g) => globToRegExp(g).test(file))) return errors;
	const skip = (ignore.patterns ?? []).map((p) => new RegExp(p));
	const dirs = [''];
	const parts = posix
		.dirname(file)
		.split('/')
		.filter((p) => p && p !== '.');
	for (let i = 1; i <= parts.length; i++) dirs.push(parts.slice(0, i).join('/'));
	for (const row of scanLines(text)) {
		if (row.fenced) continue;
		for (const m of row.raw.matchAll(/(`+)([^`\n]+?)\1/g)) {
			let token = m[2]
				.trim()
				.replace(/[.,;:]+$/, '')
				.replace(/:\d+(?:-\d+)?$/, '');
			token = token.replace(/#L\d+.*$/, '').replace(/\/+$/, '');
			const t = PATH_TOKEN.exec(token);
			if (!t || token.includes('..') || /\.\.\./.test(m[2])) continue;
			if (t[1].startsWith('packages/') && !file.startsWith('viewers/')) continue;
			// Build output, local files and a user's own project (`src/main.ts`) are not repository paths.
			if (BUILD_OUTPUT.test(t[1]) || (/^src\/[^/]+$/.test(t[1]) && file.includes('/'))) continue;
			if (skip.some((r) => r.test(t[1]))) continue;
			let rel = t[1].replace(/<[^>/]*>/g, '*');
			const brace = /\{([^{}]*)\}/.exec(rel);
			const variants = brace ? brace[1].split(',').map((o) => rel.replace(brace[0], o)) : [rel];
			const ok = variants.every((v) => dirs.some((d) => globExists(root, posix.join(d, v))));
			if (!ok) {
				errors.push({
					file,
					line: row.n,
					check: 'path',
					message: `Path \`${token}\` does not exist`,
				});
			}
		}
	}
	return errors;
}

export function checkTerms(file, text, rules) {
	const errors = [];
	const compiled = rules.map((r) => ({
		...r,
		re: new RegExp(r.pattern, r.flags ?? ''),
		allowed: (r.allow ?? []).map(globToRegExp),
	}));
	for (const rule of compiled) {
		if (rule.allowed.some((a) => a.test(file))) continue;
		text.split(/\r?\n/).forEach((line, i) => {
			if (/em-dash: intended\s*(?:-->)?\s*$/.test(line)) return;
			if (rule.re.test(line)) {
				const entry = { file, line: i + 1, check: `term:${rule.id}`, message: rule.message };
				errors.push(rule.severity === 'warning' ? { ...entry, warning: true } : entry);
			}
		});
	}
	return errors;
}

function walkMd(root, dir, out = []) {
	for (const name of readdirSync(join(root, dir))) {
		if (name === 'node_modules' || name.startsWith('.')) continue;
		const rel = dir ? `${dir}/${name}` : name;
		if (isDir(join(root, rel))) walkMd(root, rel, out);
		else if (name.endsWith('.md')) out.push(rel);
	}
	return out;
}

/** Sidebar/nav `link:` values of one docs site, and the pages no link reaches. */
export function checkSidebars(root, docRoot, files, generated = []) {
	const errors = [];
	const warnings = [];
	const configs = files.filter(
		(f) =>
			f.startsWith(`${docRoot}/.vitepress/`) &&
			/\.(ts|mts|js|mjs)$/.test(f) &&
			!f.includes('/theme/'),
	);
	const reached = new Set();
	for (const cfg of configs) {
		const text = readFileSync(join(root, cfg), 'utf8');
		text.split(/\r?\n/).forEach((line, i) => {
			for (const m of line.matchAll(/\blink:\s*(['"`])([^'"`]*)\1/g)) {
				const target = m[2].split('#')[0];
				if (/^[a-z][a-z0-9+.-]*:/i.test(target) || target.startsWith('//') || target === '') return;
				if (/\$\{/.test(target)) return;
				const hit = resolveTarget(root, `${docRoot}/index.md`, target, generated);
				if (hit === null) {
					errors.push({
						file: cfg,
						line: i + 1,
						check: 'sidebar',
						message: `Dead sidebar/nav link \`${m[2]}\``,
					});
				} else if (hit !== 'external') reached.add(hit.replace(/\.html$/, ''));
			}
		});
	}
	const pages = files.filter((f) => f.startsWith(`${docRoot}/`) && f.endsWith('.md'));
	for (const page of pages) {
		const rel = page.slice(docRoot.length + 1);
		const first = rel.split('/')[0];
		if (LOCALE.test(first) && rel.includes('/')) continue;
		if (first.startsWith('.') || first === 'public' || rel === 'index.md') continue;
		if (!reached.has(page) && !reached.has(page.replace(/\.md$/, ''))) {
			warnings.push({
				file: page,
				line: 1,
				check: 'orphan',
				message: 'English page is in no sidebar or nav',
			});
		}
	}
	return { errors, warnings };
}

/** Newest commit time per file, from one pass over the history. */
export function lastCommitTimes(root) {
	const times = new Map();
	const out = execFileSync(
		'git',
		['log', '--format=@%ct', '--name-only', '--no-renames', '--', '*.md'],
		{ cwd: root, encoding: 'utf8', maxBuffer: 1 << 28 },
	);
	let t = 0;
	for (const line of out.split('\n')) {
		if (line.startsWith('@')) t = Number(line.slice(1));
		else if (line && !times.has(line)) times.set(line, t);
	}
	return times;
}

/** Per docs site: English pages missing from, or newer than, each locale. */
export function checkTranslations(root, docRoot, files, times) {
	const pages = files.filter((f) => f.startsWith(`${docRoot}/`) && f.endsWith('.md'));
	const english = [];
	const locales = new Set();
	for (const p of pages) {
		const rel = p.slice(docRoot.length + 1);
		const first = rel.split('/')[0];
		if (LOCALE.test(first) && rel.includes('/')) locales.add(first);
		else if (!first.startsWith('.') && first !== 'public') english.push(rel);
	}
	const result = [];
	for (const locale of [...locales].sort()) {
		const missing = [];
		const stale = [];
		for (const rel of english) {
			const tr = `${docRoot}/${locale}/${rel}`;
			if (!pages.includes(tr)) {
				missing.push(rel);
				continue;
			}
			const en = times.get(`${docRoot}/${rel}`);
			const loc = times.get(tr);
			if (en !== undefined && loc !== undefined && en > loc) stale.push(rel);
		}
		result.push({ docRoot, locale, total: english.length, missing, stale });
	}
	return result;
}

export function runChecks({ root, files, config, times }) {
	const errors = [];
	const warnings = [];
	const anchorCache = new Map();
	const generated = (config.generatedRoutes ?? []).map((r) => new RegExp(r));
	const md = files.filter((f) => f.endsWith('.md') && !SKIP_FILE.test(f));
	for (const file of md) {
		const text = readFileSync(join(root, file), 'utf8');
		const l = checkLinks(root, file, text, anchorCache, generated);
		errors.push(...l.errors, ...checkCodePaths(root, file, text, config.pathIgnore));
		warnings.push(...l.warnings);
		for (const t of checkTerms(file, text, config.rules ?? [])) {
			(t.warning ? warnings : errors).push(t);
		}
	}
	const sites = [...new Set(md.map(docRootOf).filter(Boolean))].sort();
	const translations = [];
	for (const site of sites) {
		const s = checkSidebars(
			root,
			site,
			files.filter((f) => !SKIP_FILE.test(f) || f.endsWith('.ts')),
			generated,
		);
		errors.push(...s.errors);
		warnings.push(...s.warnings);
		if (times) translations.push(...checkTranslations(root, site, md, times));
	}
	return { errors, warnings, translations };
}

function section(title, items, limit = 200) {
	if (items.length === 0) return '';
	const rows = items
		.slice(0, limit)
		.map((e) => `- \`${e.file}:${e.line}\` [${e.check}] ${e.message}`);
	if (items.length > limit) rows.push(`- ... and ${items.length - limit} more`);
	return `### ${title} (${items.length})\n\n${rows.join('\n')}\n\n`;
}

export function renderReport({ errors, warnings, translations }) {
	const by = (list, kind) => list.filter((e) => e.check.split(':')[0] === kind);
	let out = `## Documentation health\n\n${errors.length} error(s), ${warnings.length} warning(s).\n\n`;
	out += section('Dead links', by(errors, 'link'));
	out += section('Missing repository paths', by(errors, 'path'));
	out += section('Stale terms', by(errors, 'term'));
	out += section('Dead sidebar and nav links', by(errors, 'sidebar'));
	out += section('Pages in no sidebar', by(warnings, 'orphan'));
	out += section('Dead anchors', by(warnings, 'anchor'));
	out += section('Stale terms (warnings)', by(warnings, 'term'));
	if (translations.length > 0) {
		out +=
			'### Translation lag\n\n| Site | Locale | English pages | Missing | Older than English |\n| --- | --- | ---: | ---: | ---: |\n';
		for (const t of translations) {
			out += `| ${t.docRoot} | ${t.locale} | ${t.total} | ${t.missing.length} | ${t.stale.length} |\n`;
		}
		out += '\n';
		for (const t of translations) {
			if (t.stale.length === 0) continue;
			const shown = t.stale
				.slice(0, 25)
				.map((p) => `\`${p}\``)
				.join(', ');
			out += `<details><summary>${t.docRoot} ${t.locale}: older than English</summary>\n\n${shown}${t.stale.length > 25 ? `, and ${t.stale.length - 25} more` : ''}\n\n</details>\n\n`;
		}
	}
	return out;
}

function main() {
	const args = process.argv.slice(2);
	const opt = (name) => {
		const i = args.indexOf(name);
		return i >= 0 ? args[i + 1] : undefined;
	};
	const root = resolve(opt('--root') ?? join(dirname(fileURLToPath(import.meta.url)), '..'));
	const config = JSON.parse(readFileSync(join(root, 'scripts/check-docs.config.json'), 'utf8'));
	let files;
	try {
		files = execFileSync('git', ['ls-files', '*.md', 'viewers/*/docs/.vitepress/**'], {
			cwd: root,
			encoding: 'utf8',
			maxBuffer: 1 << 28,
		})
			.split('\n')
			.filter(Boolean);
	} catch {
		files = walkMd(root, '');
	}
	files = files.filter((f) => existsSync(join(root, f)));
	let times;
	try {
		times = lastCommitTimes(root);
	} catch {
		times = undefined;
	}
	const result = runChecks({ root, files, config, times });
	const report = renderReport(result);
	const out = opt('--out');
	if (out) writeFileSync(out, report);
	else process.stdout.write(report);
	console.error(
		`docs check: ${result.errors.length} error(s), ${result.warnings.length} warning(s)`,
	);
	if (result.errors.length > 0 && !args.includes('--warn-only')) process.exit(1);
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) main();
