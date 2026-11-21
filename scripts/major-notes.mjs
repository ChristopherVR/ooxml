#!/usr/bin/env node
/**
 * The hand-written notes of a coordinated major release, read from `.release-notes/<name>/`.
 *
 * A release name (`2.0`) is not a package version: every package moves to its own next major.
 * The folder holds `overview.md`, a few paragraphs every package's notes start with, and one
 * fragment per merged piece of work. A fragment has `## <kind>` sections (KINDS) of bullets, and
 * every bullet starts with the packages it affects, then a colon:
 *
 *   - `ooxml-core/teams`, `openteams-server` and OpenTeams viewers: the new chat ids are ...
 *
 * A name in backticks is an npm name from the release planner's table or a subpath of one
 * (`ooxml-core/opc`). A name without backticks is one of the groups in `groups()`. `Office suite`
 * names the hosted app, which is not on npm: its bullets reach no package changelog.
 *
 * When a package is released with a major bump and its changelog does not carry these notes yet
 * (a hand-written section at the top, see crafted-changelogs.mjs, carries them),
 * the release workflow (`--plan`) writes the package's share of them under the new version's
 * heading, above the git-cliff commit list, which moves under `### Commits`. `release-notes.mjs`
 * gives the GitHub release body the same shape. A bullet that names no package, an unknown name,
 * or a published package with no bullet fails `--check`, `test:scripts` and the release run.
 *
 *   node scripts/major-notes.mjs --check
 *   node scripts/major-notes.mjs --plan release-plan.json
 *   node scripts/major-notes.mjs --render <package key> [name]
 */
import { readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { pendingSection } from './crafted-changelogs.mjs';
import { PACKAGES, presentPackages } from './release-plan.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
export const NOTES_DIR = '.release-notes';
export const OVERVIEW = 'overview.md';

/** Fragment section names, in the order a package's notes list them, and their headings there. */
export const KINDS = [
	['Breaking', 'Breaking changes'],
	['Migration', 'Migration'],
	['Security', 'Security'],
	['Added', 'Added'],
	['Changed', 'Changed'],
	['Fixed', 'Fixed'],
	['Performance', 'Performance'],
	['Accessibility', 'Accessibility'],
];
export const COMMITS_HEADING = 'Commits';

/** The comment that marks a changelog section as carrying the notes of release `name`. */
export const marker = (name) => `<!-- release-notes: ${name} -->`;

const BINDING_DIR =
	/^viewers\/(?<viewer>[^/]+)\/packages\/(react|vue|angular|svelte|solid|vanilla|cli)$/u;
const MCP_DIR = /^(?:mcp|viewers\/[^/]+\/mcp|viewers\/pptx\/packages\/tools)$/u;

/** The group names a lead-in may use, each with the npm names it stands for. */
export function groups(table = PACKAGES) {
	const metas = Object.values(table);
	const bindings = (viewer) =>
		metas.filter((m) => BINDING_DIR.exec(m.dir)?.groups.viewer === viewer).map((m) => m.npm);
	return new Map([
		['All packages', metas.map((m) => m.npm)],
		['PowerPoint viewers', bindings('pptx')],
		['Word viewers', bindings('docx')],
		['Excel viewers', bindings('xlsx')],
		['Visio viewers', bindings('visio')],
		['OpenTeams viewers', bindings('teams')],
		['MCP packages', metas.filter((m) => MCP_DIR.test(m.dir)).map((m) => m.npm)],
		['Office suite', []],
	]);
}

/** The npm names one lead-in item stands for, `{ unknown }` for an unknown name, or null for prose. */
function resolveItem(item, table, named) {
	const code = /^`(?<name>[^`\s]+)`$/u.exec(item)?.groups.name;
	if (code === undefined) return named.get(item) ?? null;
	const npm = Object.values(table)
		.map((m) => m.npm)
		.find((name) => code === name || code.startsWith(`${name}/`));
	return npm ? [npm] : { unknown: code };
}

const HELP =
	'start each bullet with the packages it affects and a colon, such as "- `ooxml-core`, PowerPoint viewers: ..." (see .release-notes/README.md)';

/** The entries and problems of one fragment. `file` is the path problems name. */
export function parseFragment(text, file, table = PACKAGES) {
	const named = groups(table);
	const kinds = new Map(KINDS);
	const entries = [];
	const problems = [];
	let kind = null;
	let last = null;
	text
		.replace(/\r\n/gu, '\n')
		.split('\n')
		.forEach((line, index) => {
			const at = `${file}:${index + 1}`;
			if (line.trim() === '') return;
			if (line.startsWith('## ')) {
				kind = line.slice(3).trim();
				last = null;
				if (!kinds.has(kind)) {
					problems.push(`${at}: unknown section "${kind}" (use ${[...kinds.keys()].join(', ')})`);
				}
			} else if (line.startsWith('- ')) {
				last = null;
				if (!kind) return problems.push(`${at}: bullet before any "## <section>" heading`);
				const lead = /^- (?<lead>.+?): \S/u.exec(line)?.groups.lead;
				const items = lead ? lead.split(/,\s+and\s+|,\s+|\s+and\s+/u) : [];
				const resolved = items.map((item) => resolveItem(item, table, named));
				if (!lead || resolved.includes(null))
					return problems.push(`${at}: names no package; ${HELP}`);
				const unknown = resolved.filter((r) => r.unknown).map((r) => `\`${r.unknown}\``);
				if (unknown.length) {
					return problems.push(
						`${at}: unknown package ${unknown.join(', ')}; use an npm name from scripts/release-plan.mjs, a subpath of one, or a group (${[...named.keys()].join(', ')})`,
					);
				}
				last = { kind, text: line.slice(2), packages: [...new Set(resolved.flat())], at };
				entries.push(last);
			} else if (/^\s+\S/u.test(line) && last) {
				last.text += `\n${line}`;
			} else {
				problems.push(`${at}: expected a "## <section>" heading or a "- " bullet`);
			}
		});
	return { entries, problems };
}

const natural = (a, b) => a.localeCompare(b, 'en', { numeric: true });

function readOptionalText(path) {
	try {
		return readFileSync(path, 'utf8');
	} catch (error) {
		if (error.code === 'ENOENT') return '';
		throw error;
	}
}

/** The notes of release `name`: overview, entries and every problem, coverage included. */
export function loadNotes(root, name, table = PACKAGES) {
	const dir = join(root, NOTES_DIR, name);
	const rel = `${NOTES_DIR}/${name}`;
	const problems = [];
	const overviewPath = join(dir, OVERVIEW);
	const overview = readOptionalText(overviewPath).trim();
	if (!overview) problems.push(`${rel}/${OVERVIEW}: missing or empty`);
	const entries = [];
	for (const file of readdirSync(dir)
		.filter((f) => f.endsWith('.md') && f !== OVERVIEW)
		.sort(natural)) {
		const parsed = parseFragment(readFileSync(join(dir, file), 'utf8'), `${rel}/${file}`, table);
		entries.push(...parsed.entries);
		problems.push(...parsed.problems);
	}
	for (const meta of Object.values(presentPackages(root, table))) {
		if (!entries.some((entry) => entry.packages.includes(meta.npm))) {
			problems.push(`${rel}: no entry names \`${meta.npm}\`; every published package needs one`);
		}
	}
	return { name, overview, entries, problems };
}

/** `loadNotes`, throwing when the notes have a problem, so a release never ships them broken. */
export function requireNotes(root, name, table = PACKAGES) {
	const notes = loadNotes(root, name, table);
	if (notes.problems.length) {
		throw new Error(`Release notes ${name} are not ready:\n${notes.problems.join('\n')}`);
	}
	return notes;
}

/** The release notes names under `root`, oldest first. */
export function noteNames(root) {
	const dir = join(root, NOTES_DIR);
	let entries;
	try {
		entries = readdirSync(dir, { withFileTypes: true });
	} catch (error) {
		if (error.code === 'ENOENT') return [];
		throw error;
	}
	return entries
		.filter((d) => d.isDirectory())
		.map((d) => d.name)
		.sort(natural);
}

/** One package's share of the notes: marker, overview, then its bullets by kind. */
export function renderPackageNotes(notes, npm) {
	const lines = [marker(notes.name), '', notes.overview, ''];
	for (const [kind, heading] of KINDS) {
		const own = notes.entries.filter((e) => e.kind === kind && e.packages.includes(npm));
		if (own.length) lines.push(`### ${heading}`, '', ...own.map((e) => `- ${e.text}`), '');
	}
	return lines.join('\n').trimEnd();
}

/** A version section's body: the notes `blocks`, then the git-cliff `body` under `### Commits`. */
export function craftedSection(blocks, body) {
	const parts = [...blocks];
	const commits = body.trim();
	if (/^\s*[-*]\s/mu.test(commits)) {
		parts.push(`### ${COMMITS_HEADING}`, commits.replace(/^### /gmu, '#### '));
	}
	return `${parts.join('\n\n')}\n`;
}

/** `text` split around the section of `version`: lines before, its heading, its body, lines after. */
function splitSection(text, version) {
	const lines = text.replace(/\r\n/gu, '\n').split('\n');
	const start = lines.findIndex((line) => line.startsWith(`## [${version}]`));
	if (start < 0) return null;
	const next = lines.findIndex((line, i) => i > start && line.startsWith('## '));
	const end = next < 0 ? lines.length : next;
	return {
		before: lines.slice(0, start),
		heading: lines[start],
		body: lines.slice(start + 1, end).join('\n'),
		after: lines.slice(end),
	};
}

/**
 * The notes names `pkg` (a release plan entry) must carry: a major release not carrying them yet.
 * A changelog with a hand-written section waiting for its date (crafted-changelogs.mjs) needs none.
 */
export function pendingNotes(root, pkg, changelog = '') {
	if (!pkg.release || pkg.bump !== 'major' || pendingSection(changelog)) return [];
	const parts = splitSection(changelog, pkg.version);
	const elsewhere = parts ? [...parts.before, ...parts.after].join('\n') : changelog;
	return noteNames(root).filter((name) => !elsewhere.includes(marker(name)));
}

/** `text` with `blocks` written into the section of `version`. Running it twice changes nothing. */
export function insertIntoChangelog(text, version, names, blocks) {
	const parts = splitSection(text, version);
	if (!parts) throw new Error(`no "## [${version}]" section to put the release notes in`);
	const fresh = names
		.map((name, i) => [name, blocks[i]])
		.filter(([name]) => !parts.body.includes(marker(name)));
	if (!fresh.length) return text;
	const section = `${parts.heading}\n\n${craftedSection(
		fresh.map(([, block]) => block),
		parts.body,
	)}`;
	return [...parts.before, section, ...parts.after].join('\n');
}

/** A GitHub release body (`## [version]` heading optional) with the notes `blocks` above its commits. */
export function releaseBody(notes, blocks) {
	if (!blocks.length) return notes;
	const [first, ...rest] = notes.split('\n');
	const heading = first.startsWith('## ') ? first : null;
	const body = craftedSection(blocks, heading ? rest.join('\n') : notes);
	return heading ? `${heading}\n\n${body}` : body;
}

/** The notes blocks of `pkg`, read from `root` and checked. */
export function packageBlocks(root, pkg, changelog, table = PACKAGES) {
	const names = pendingNotes(root, pkg, changelog);
	return {
		names,
		blocks: names.map((name) => renderPackageNotes(requireNotes(root, name, table), pkg.npm)),
	};
}

function main() {
	const args = process.argv.slice(2);
	if (args[0] === '--check') {
		const problems = noteNames(ROOT).flatMap((name) => loadNotes(ROOT, name).problems);
		for (const problem of problems) console.error(problem);
		process.exit(problems.length ? 1 : 0);
	} else if (args[0] === '--plan' && args[1]) {
		const plan = JSON.parse(readFileSync(args[1], 'utf8'));
		for (const key of plan.order) {
			const pkg = plan.packages[key];
			const path = join(ROOT, pkg.changelog);
			const text = readOptionalText(path);
			const { names, blocks } = packageBlocks(ROOT, pkg, text);
			if (!names.length) continue;
			const next = insertIntoChangelog(text, pkg.version, names, blocks);
			if (next !== text) {
				writeFileSync(path, next);
				console.log(`${pkg.changelog}: release notes ${names.join(', ')}`);
			}
		}
	} else if (args[0] === '--render' && PACKAGES[args[1]]) {
		const name = args[2] ?? noteNames(ROOT).at(-1);
		console.log(renderPackageNotes(requireNotes(ROOT, name), PACKAGES[args[1]].npm));
	} else {
		console.error(
			'usage: node scripts/major-notes.mjs --check | --plan <release-plan.json> | --render <key> [name]',
		);
		process.exit(2);
	}
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
