/**
 * The root changelog of each viewer (`viewers/<name>/CHANGELOG.md`): one dated section per release
 * day, naming every package of that viewer released that day and merging their entries.
 *
 * The release workflow writes each package's own changelog with git-cliff; this script then
 * rebuilds the viewer's section for the day from those package sections, so the two never
 * disagree. A day's section is deterministic (every package section with that date), which makes
 * the script idempotent and lets a second release run on the same day extend the day's section
 * instead of adding another. Sections for other days are never touched (changelogs here are
 * prepend-only, see check-changelog-sections.mjs).
 *
 * Usage:
 *   node scripts/viewer-changelogs.mjs --plan release-plan.json   (release workflow, after git-cliff)
 *   node scripts/viewer-changelogs.mjs --backfill                 (every package day missing a section)
 * Both print the files they changed.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { COMMITS_HEADING, KINDS } from './major-notes.mjs';
import { VIEWER_PACKAGES } from './viewer-packages.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/** cliff.toml's group order; a group not listed here sorts after these, by first appearance. */
const GROUP_ORDER = [
	'Features',
	'Bug Fixes',
	'Performance',
	'Refactor',
	'Documentation',
	'Testing',
	'Build & CI',
	'Styling',
	'Dependencies',
	'Chores',
	'Reverts',
	'Security',
	'Other',
];
const NOTES_ORDER = KINDS.map(([, heading]) => heading);

export const HEADER = [
	'# Changelog',
	'',
	"All notable changes to this viewer's packages, newest first. Each dated section names the",
	'packages released that day and merges their entries; the full history of each package is in its',
	'own `CHANGELOG.md` (`packages/<name>/`, `mcp/`, `server/`). The release workflow writes these',
	'sections (`scripts/viewer-changelogs.mjs`); do not edit them by hand.',
	'',
].join('\n');

const PACKAGE_HEADING_RE = /^## \[(?<version>[^\]]+)\]\([^)]*\) - (?<date>\d{4}-\d{2}-\d{2})\s*$/u;
const DATED_HEADING_RE = /^## (?<date>\d{4}-\d{2}-\d{2})\s*$/u;

/** The released sections of one package changelog: version, date and entries grouped by heading. */
export function parsePackageChangelog(text) {
	const sections = [];
	let section = null;
	let group = null;
	for (const line of text.replace(/\r\n/gu, '\n').split('\n')) {
		const heading = PACKAGE_HEADING_RE.exec(line);
		if (heading) {
			section = { ...heading.groups, groups: [] };
			group = null;
			sections.push(section);
		} else if (line.startsWith('## ')) {
			section = null;
		} else if (section && line.startsWith('### ')) {
			group = { heading: line.slice(4).trim(), entries: [] };
			section.groups.push(group);
		} else if (group && line.startsWith('- ')) {
			group.entries.push(line);
		} else if (group && /^\s+\S/u.test(line) && group.entries.length) {
			group.entries[group.entries.length - 1] += `\n${line}`;
		}
	}
	return sections;
}

/** `sections` merged into one group list in cliff order, each entry once. */
export function mergeGroups(sections) {
	const byHeading = new Map();
	for (const section of sections)
		for (const group of section.groups) {
			const entries = byHeading.get(group.heading) ?? [];
			for (const entry of group.entries) if (!entries.includes(entry)) entries.push(entry);
			byHeading.set(group.heading, entries);
		}
	// A day whose releases carry hand-written notes (major-notes.mjs) lists them in their own order,
	// and the commits, which sit under `### Commits` there, last.
	const notes = byHeading.has(COMMITS_HEADING);
	const order = notes ? NOTES_ORDER : GROUP_ORDER;
	const rank = (heading) => {
		if (notes && heading === COMMITS_HEADING) return Infinity;
		const index = order.indexOf(heading);
		return index < 0 ? order.length : index;
	};
	return [...byHeading]
		.filter(([, entries]) => entries.length)
		.sort(([a], [b]) => rank(a) - rank(b))
		.map(([heading, entries]) => ({ heading, entries }));
}

/** One day's viewer section from the package releases of that day (`{ npm, version, groups }`). */
export function renderDaySection(date, releases) {
	const lines = [
		`## ${date}`,
		'',
		`_Releases: ${releases.map((r) => `${r.npm}@${r.version}`).join(', ')}_`,
		'',
	];
	for (const { heading, entries } of mergeGroups(releases))
		lines.push(`### ${heading}`, '', ...entries, '');
	return lines.join('\n');
}

/**
 * `text` with the section for `section`'s date replaced, or inserted in date order (newest first)
 * when there is none, and the header brought up to date. Other sections are kept as they are.
 */
export function upsertDaySection(text, date, section) {
	const normalised = text.replace(/\r\n/gu, '\n');
	const first = normalised.search(/^## /mu);
	const body = first < 0 ? '' : normalised.slice(first);
	const parts = body ? body.split(/^(?=## )/mu) : [];
	const index = parts.findIndex(
		(part) => DATED_HEADING_RE.exec(part.split('\n')[0])?.groups.date === date,
	);
	const block = `${section.trimEnd()}\n\n`;
	if (index >= 0) {
		parts[index] = block;
	} else {
		const at = parts.findIndex((part) => {
			const found = DATED_HEADING_RE.exec(part.split('\n')[0]);
			return found ? found.groups.date < date : false;
		});
		parts.splice(at < 0 ? parts.length : at, 0, block);
	}
	return `${HEADER}\n${parts.join('').trimEnd()}\n`;
}

/** The viewers and, for each, the changelog paths and npm names of its published packages. */
export function viewerPackageTable(table = VIEWER_PACKAGES) {
	const viewers = new Map();
	for (const meta of Object.values(table)) {
		const name = /^viewers\/([^/]+)\//u.exec(meta.dir)?.[1];
		if (!name) continue;
		const list = viewers.get(name) ?? [];
		list.push({ npm: meta.npm, changelog: `${meta.dir}/CHANGELOG.md` });
		viewers.set(name, list);
	}
	return viewers;
}

/** Every package release of one viewer, read from the package changelogs under `root`. */
function viewerReleases(root, packages) {
	return packages.flatMap(({ npm, changelog }) => {
		const path = join(root, changelog);
		if (!existsSync(path)) return [];
		return parsePackageChangelog(readFileSync(path, 'utf8')).map((s) => ({ npm, ...s }));
	});
}

/** The dates already covered by a viewer's root changelog text. */
function coveredDates(text) {
	return new Set(
		text
			.replace(/\r\n/gu, '\n')
			.split('\n')
			.map((line) => DATED_HEADING_RE.exec(line)?.groups.date)
			.filter(Boolean),
	);
}

/**
 * Write the sections for `datesFor(viewer, releases, existing)` into each viewer's root changelog.
 * Returns the paths it changed.
 */
export function updateViewerChangelogs(root, datesFor, table = VIEWER_PACKAGES) {
	const changed = [];
	for (const [viewer, packages] of viewerPackageTable(table)) {
		const releases = viewerReleases(root, packages);
		const path = join(root, 'viewers', viewer, 'CHANGELOG.md');
		const before = existsSync(path) ? readFileSync(path, 'utf8') : `${HEADER}\n`;
		let text = before;
		for (const date of [...datesFor(viewer, releases, before)].sort()) {
			const day = releases.filter((r) => r.date === date);
			if (day.length) text = upsertDaySection(text, date, renderDaySection(date, day));
		}
		if (text !== before) {
			writeFileSync(path, text);
			changed.push(`viewers/${viewer}/CHANGELOG.md`);
		}
	}
	return changed;
}

/** Release mode: the day of each viewer package `plan` released. */
export function planDates(plan) {
	const released = Object.values(plan.packages).filter((p) => p.release);
	return (viewer, releases) =>
		new Set(
			released
				.filter((p) => p.dir.startsWith(`viewers/${viewer}/`))
				.flatMap((p) => releases.filter((r) => r.npm === p.npm && r.version === p.version))
				.map((r) => r.date),
		);
}

/** Backfill mode: every package release day the viewer changelog has no section for. */
export const missingDates = (_viewer, releases, existing) => {
	const covered = coveredDates(existing);
	return new Set(releases.map((r) => r.date).filter((date) => !covered.has(date)));
};

function main() {
	const args = process.argv.slice(2);
	const planAt = args.indexOf('--plan');
	let datesFor;
	if (planAt >= 0 && args[planAt + 1]) {
		datesFor = planDates(JSON.parse(readFileSync(args[planAt + 1], 'utf8')));
	} else if (args.includes('--backfill')) {
		datesFor = missingDates;
	} else {
		console.error(
			'Usage: node scripts/viewer-changelogs.mjs --plan <release-plan.json> | --backfill',
		);
		process.exit(2);
	}
	for (const path of updateViewerChangelogs(ROOT, datesFor)) console.log(path);
}

if (process.argv[1]?.endsWith('viewer-changelogs.mjs')) main();
