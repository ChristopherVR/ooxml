#!/usr/bin/env node
/**
 * Hand-written changelog sections: a release section written before the release, which the release
 * workflow publishes as it stands instead of generating one with git-cliff.
 *
 * A package changelog carries one as its top section, with `Unreleased` where the date goes:
 *
 *   ## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@2.0.0) - Unreleased
 *
 *   <!-- release-notes: 2.0 -->
 *
 * A viewer root changelog (`viewers/<name>/CHANGELOG.md`) carries `## Unreleased` with the usual
 * `_Releases: <npm>@<version>, ..._` line. When the plan releases the package (or, for a viewer
 * section, a package its release line names), the workflow skips git-cliff and the major-notes
 * insertion for it and replaces `Unreleased` with the release date (UTC). A section whose version
 * or tag does not match the plan fails the run before anything is built or tagged.
 *
 *   node scripts/crafted-changelogs.mjs --check <release-plan.json>    problems, exit 1 on any
 *   node scripts/crafted-changelogs.mjs --pending <release-plan.json>  keys skipping git-cliff
 *   node scripts/crafted-changelogs.mjs --date <release-plan.json>     date the package sections
 *
 * The viewer sections are dated by `viewer-changelogs.mjs --plan`. Dependency-free: it runs only
 * `node:` builtins.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

export const PENDING = 'Unreleased';
export const RELEASES_URL = 'https://github.com/ChristopherVR/ooxml/releases/tag/';

/** The tag URL a package section heading links to. */
export const releaseUrl = (npm, version) => `${RELEASES_URL}${npm}@${version}`;

const PACKAGE_PENDING_RE = /^## \[(?<version>[^\]]+)\]\((?<url>[^)\s]*)\) - Unreleased\s*$/u;
const VIEWER_PENDING_RE = /^## Unreleased\s*$/u;
const MARKER_RE = /^<!-- release-notes: [^>]*-->\s*$/u;
const RELEASES_RE = /^_Releases:(?<list>.*)_\s*$/u;

/** The `## ` sections of `text` (the file header is not one): heading, line index, body lines. */
export function sections(text) {
	const lines = text.replace(/\r\n/gu, '\n').split('\n');
	const found = [];
	lines.forEach((line, index) => {
		if (line.startsWith('## ')) found.push({ heading: line, index, body: [] });
		else found.at(-1)?.body.push(line);
	});
	return found;
}

/** `package`, `viewer` or null: which pending form a heading line has. */
export function pendingKind(heading) {
	if (PACKAGE_PENDING_RE.test(heading)) return 'package';
	if (VIEWER_PENDING_RE.test(heading)) return 'viewer';
	return null;
}

/** The pending package section at the top of `text`: `{ version, url, body }`, or null. */
export function pendingSection(text) {
	const [top] = sections(text);
	const match = top && PACKAGE_PENDING_RE.exec(top.heading);
	return match ? { ...match.groups, body: top.body } : null;
}

/** The pending viewer section at the top of `text`: `{ releases, body }`, or null. */
export function pendingViewerSection(text) {
	const [top] = sections(text);
	if (!top || !VIEWER_PENDING_RE.test(top.heading)) return null;
	return { releases: releaseList(top.body), body: top.body };
}

/** The `npm@version` names of a section's `_Releases: ..._` line, or null without one. */
export function releaseList(body) {
	for (const line of body) {
		const match = RELEASES_RE.exec(line.trim());
		if (match)
			return match.groups.list
				.split(',')
				.map((name) => name.trim())
				.filter(Boolean);
	}
	return null;
}

/** A section body as published notes: no release-notes marker, no blank lines at either end. */
export function notesBody(body) {
	return body
		.filter((line) => !MARKER_RE.test(line.trim()))
		.join('\n')
		.trim();
}

/** Why the pending section of `text` cannot be published as `pkg`'s release; [] when it can. */
export function packageProblems(pkg, text, path = pkg.changelog) {
	const pending = pendingSection(text);
	if (!pending) return [];
	const problems = [];
	if (pending.version !== pkg.version) {
		problems.push(
			`${path}: the hand-written section is for ${pending.version}, but this release plans ${pkg.npm}@${pkg.version}; change its heading to ## [${pkg.version}](${releaseUrl(pkg.npm, pkg.version)}) - ${PENDING}`,
		);
	} else if (pending.url !== releaseUrl(pkg.npm, pkg.version)) {
		problems.push(
			`${path}: the hand-written section links ${pending.url}, expected ${releaseUrl(pkg.npm, pkg.version)}`,
		);
	}
	if (!notesBody(pending.body)) problems.push(`${path}: the hand-written section is empty`);
	return problems;
}

/**
 * Why the pending section of a viewer changelog cannot be published with this plan; [] when it
 * can, or when the plan releases none of the packages it names. `released` maps the npm name of
 * every package of the viewer the plan releases to its version.
 */
export function viewerProblems(released, text, path) {
	const pending = pendingViewerSection(text);
	if (!pending) return [];
	if (pending.releases === null || pending.releases.length === 0) {
		return [`${path}: the ${PENDING} section has no \`_Releases: <npm>@<version>, ..._\` line`];
	}
	const named = pending.releases.map((entry) => /^(?<npm>.+)@(?<version>[^@]+)$/u.exec(entry));
	if (!named.some((m) => m && released.has(m.groups.npm))) return [];
	return pending.releases.flatMap((entry, i) => {
		const m = named[i];
		if (!m) return [`${path}: "${entry}" in the ${PENDING} section is not <npm>@<version>`];
		const planned = released.get(m.groups.npm);
		if (planned === m.groups.version) return [];
		return [
			planned
				? `${path}: the ${PENDING} section names ${entry}, but this release plans ${m.groups.npm}@${planned}`
				: `${path}: the ${PENDING} section names ${entry}, which this release does not release`,
		];
	});
}

/** `text` with its top pending heading (package or viewer form) dated `date`. */
export function datePending(text, date) {
	const lines = text.replace(/\r\n/gu, '\n').split('\n');
	const [top] = sections(text);
	if (!top) return text;
	const kind = pendingKind(top.heading);
	if (kind === 'package')
		lines[top.index] = top.heading.replace(/ - Unreleased\s*$/u, ` - ${date}`);
	else if (kind === 'viewer') lines[top.index] = `## ${date}`;
	else return text;
	return lines.join('\n');
}

/** Today's date in UTC, `YYYY-MM-DD`, as git-cliff writes it. */
export const today = () => new Date().toISOString().slice(0, 10);

/**
 * The GitHub release body of `pkg` from a hand-written section, or null. `texts` are versions of
 * its changelog, tried in order: the working tree (before or after dating) and the committed file.
 */
export function craftedNotes(pkg, texts) {
	for (const text of texts) {
		const pending = text ? pendingSection(text) : null;
		if (pending && pending.version === pkg.version) return notesBody(pending.body) || null;
	}
	return null;
}

const read = (root, path) => {
	const file = join(root, path);
	return existsSync(file) ? readFileSync(file, 'utf8') : '';
};

const releasedKeys = (plan) => plan.order.filter((key) => plan.packages[key].release);

/** The npm-to-version map of the packages `plan` releases under `viewers/<viewer>/`. */
export function viewerReleases(plan, viewer) {
	return new Map(
		Object.values(plan.packages)
			.filter((p) => p.release && p.dir.startsWith(`viewers/${viewer}/`))
			.map((p) => [p.npm, p.version]),
	);
}

/** Every problem with the hand-written sections `plan` would publish under `root`. */
export function planProblems(root, plan) {
	const problems = releasedKeys(plan).flatMap((key) => {
		const pkg = plan.packages[key];
		return packageProblems(pkg, read(root, pkg.changelog));
	});
	const viewers = new Set(
		Object.values(plan.packages)
			.map((p) => /^viewers\/([^/]+)\//u.exec(p.dir)?.[1])
			.filter(Boolean),
	);
	for (const viewer of viewers) {
		const path = `viewers/${viewer}/CHANGELOG.md`;
		problems.push(...viewerProblems(viewerReleases(plan, viewer), read(root, path), path));
	}
	return problems;
}

/** The keys of the released packages whose changelog has a pending section. */
export function pendingKeys(root, plan) {
	return releasedKeys(plan).filter((key) =>
		pendingSection(read(root, plan.packages[key].changelog)),
	);
}

/** Date the pending section of every released package; returns the changelogs it changed. */
export function datePackages(root, plan, date = today()) {
	const problems = planProblems(root, plan);
	if (problems.length) throw new Error(problems.join('\n'));
	return pendingKeys(root, plan).map((key) => {
		const path = plan.packages[key].changelog;
		writeFileSync(join(root, path), datePending(read(root, path), date));
		return path;
	});
}

function main() {
	const [mode, planPath] = process.argv.slice(2);
	if (!['--check', '--pending', '--date'].includes(mode) || !planPath) {
		console.error(
			'usage: node scripts/crafted-changelogs.mjs --check | --pending | --date <release-plan.json>',
		);
		process.exit(2);
	}
	const plan = JSON.parse(readFileSync(planPath, 'utf8'));
	if (mode === '--pending') {
		console.log(pendingKeys(ROOT, plan).join(' '));
		return;
	}
	const problems = planProblems(ROOT, plan);
	for (const problem of problems) console.error(`::error::${problem}`);
	if (problems.length) process.exit(1);
	if (mode === '--date') for (const path of datePackages(ROOT, plan)) console.log(`${path}: dated`);
	else console.log(`OK: ${pendingKeys(ROOT, plan).length} hand-written section(s) match the plan.`);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
