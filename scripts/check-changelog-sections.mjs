/**
 * Adapted from ChristopherVR/pptx-viewer scripts/check-changelog-sections.mjs (Apache-2.0). Candidate for
 * centralisation: only the package tables and URLs differ between the repos.
 *
 * Guard: no changelog section may be a stub.
 *
 * Every CHANGELOG.md in this repo is PREPEND-ONLY. Old tags are culled weekly by
 * prune-releases.mjs, so a changelog can never be regenerated from tag history
 * without permanently losing the sections whose tags are gone. That makes a bad
 * section written by the release pipeline effectively permanent: the only way
 * out is a hand edit, which is exactly the operation the prepend-only rule
 * exists to avoid.
 *
 * The failure this catches actually happened. The hourly Release run bumps a
 * package whenever it changed since its own tag; when the only qualifying
 * package is a framework binding whose build then fails, the workflow reverts
 * the bump and drops it from the plan (release.yml, "Drop build-failed packages"),
 * leaving nothing to release. The root-changelog step ran anyway and prepended
 *
 *   ## 2026-08-13
 *
 *   _Releases: _
 *
 * a dated section that documents nothing, once an hour, for as long as the build
 * stayed broken. Seven accumulated in CHANGELOG.md.
 *
 * Three invariants, all cheap:
 *   1. No section anywhere declares an EMPTY release list (`_Releases: _`).
 *   2. Every dated section in the root changelog declares a release list at all.
 *   3. A hand-written section waiting for its release date (`## [<version>](<tag url>) - Unreleased`
 *      or, in a viewer root changelog, `## Unreleased`; see crafted-changelogs.mjs) is the top
 *      section, and its link is the tag of its own version. The release workflow dates only the top
 *      section, so a pending section anywhere else would be published undated.
 *
 * Usage: node scripts/check-changelog-sections.mjs [file...]
 *   With no arguments it checks every package changelog of the release planner's table, the root
 *   CHANGELOG.md, packages/<pkg>/CHANGELOG.md and each viewer's root changelog.
 */

import { readdirSync, readFileSync, existsSync } from 'node:fs';
import { dirname, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

import { pendingKind, RELEASES_URL } from './crafted-changelogs.mjs';
import { PACKAGES } from './release-plan.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/** A dated root section: `## 2026-08-13`. Per-package headings are versioned. */
const DATED_HEADING_RE = /^##\s+\d{4}-\d{2}-\d{2}\s*$/u;
/** The root changelog's release-list line, e.g. `_Releases: ooxml-core@2.3.6_`. */
const RELEASES_RE = /^_Releases:(?<list>.*)_$/u;

/**
 * Split changelog text into sections: everything from one `## ` heading up to
 * (excluding) the next one. Text above the first heading is the file header and
 * is not a section.
 */
export function parseSections(text) {
	const lines = text.split('\n');
	const sections = [];
	let current = null;
	lines.forEach((line, index) => {
		if (line.startsWith('## ')) {
			current = { heading: line.trim(), line: index + 1, body: [] };
			sections.push(current);
		} else if (current) {
			current.body.push(line);
		}
	});
	return sections;
}

/**
 * Problems with one parsed section, as plain strings. A section is a stub when
 * it announces a release run that released nothing.
 */
export function sectionProblems(section, { requireReleases }) {
	const problems = [];
	const releaseLines = section.body
		.map((l) => RELEASES_RE.exec(l.trim()))
		.filter((m) => m !== null);

	if (releaseLines.some((m) => m.groups.list.trim().length === 0)) {
		problems.push(
			'declares an empty release list (`_Releases: _`), so it documents nothing. ' +
				'A release run that released no package must not write a section at all.',
		);
	}

	if (requireReleases && DATED_HEADING_RE.test(section.heading) && releaseLines.length === 0) {
		problems.push('is a dated section with no `_Releases: ..._` line naming what it released.');
	}

	return problems;
}

const PENDING_LINK_RE = /^## \[(?<version>[^\]]+)\]\((?<url>[^)\s]*)\)/u;

/** Problems with a pending (`Unreleased`) section that is the `index`th section of its file. */
export function pendingProblems(section, index) {
	const kind = pendingKind(section.heading);
	if (!kind) return [];
	const problems = [];
	if (index > 0) {
		problems.push(
			'is waiting for its release date but is not the top section; only the top section may be ' +
				'pending, because the release workflow dates only that one.',
		);
	}
	if (kind === 'package') {
		const { version, url } = PENDING_LINK_RE.exec(section.heading).groups;
		const tag = url.startsWith(RELEASES_URL) ? url.slice(RELEASES_URL.length) : '';
		if (!tag.endsWith(`@${version}`) || tag.length <= version.length + 1) {
			problems.push(`links ${url}, which is not the tag ${RELEASES_URL}<npm>@${version}.`);
		}
	}
	return problems;
}

/** Every problem in one changelog file. `requireReleases` applies to the root. */
export function checkChangelog(text, { requireReleases = false } = {}) {
	return parseSections(text).flatMap((section, index) =>
		[...sectionProblems(section, { requireReleases }), ...pendingProblems(section, index)].map(
			(problem) => ({
				line: section.line,
				heading: section.heading,
				problem,
			}),
		),
	);
}

/**
 * Root changelog plus every packages/<pkg>/CHANGELOG.md that exists, and the same for each
 * imported viewer (its root changelog, viewers/<name>/packages/<pkg>, mcp and server).
 */
export function changelogFiles(root = ROOT) {
	const files = [];
	const rootChangelog = join(root, 'CHANGELOG.md');
	if (existsSync(rootChangelog)) {
		files.push(rootChangelog);
	}
	const scan = (packagesDir) => {
		if (!existsSync(packagesDir)) {
			return;
		}
		for (const entry of readdirSync(packagesDir, { withFileTypes: true })) {
			if (!entry.isDirectory()) {
				continue;
			}
			const candidate = join(packagesDir, entry.name, 'CHANGELOG.md');
			if (existsSync(candidate)) {
				files.push(candidate);
			}
		}
	};
	scan(join(root, 'packages'));
	const viewersDir = join(root, 'viewers');
	if (existsSync(viewersDir)) {
		for (const viewer of readdirSync(viewersDir, { withFileTypes: true })) {
			if (!viewer.isDirectory()) {
				continue;
			}
			const viewerChangelog = join(viewersDir, viewer.name, 'CHANGELOG.md');
			if (existsSync(viewerChangelog)) {
				files.push(viewerChangelog);
			}
			scan(join(viewersDir, viewer.name, 'packages'));
			for (const single of ['mcp', 'server']) {
				const candidate = join(viewersDir, viewer.name, single, 'CHANGELOG.md');
				if (existsSync(candidate)) {
					files.push(candidate);
				}
			}
		}
	}
	// Every published package, wherever it lives (src/core, src/ui, mcp, viewers/pptx/packages/tools).
	for (const meta of Object.values(PACKAGES)) {
		const candidate = join(root, meta.dir, 'CHANGELOG.md');
		if (existsSync(candidate) && !files.includes(candidate)) {
			files.push(candidate);
		}
	}
	return files;
}

function main() {
	const files = process.argv.slice(2);
	const targets = files.length > 0 ? files : changelogFiles();
	let failures = 0;

	for (const file of targets) {
		// The repository and viewer root changelogs hold dated sections that must name releases.
		const path = relative(ROOT, file).split(sep).join('/');
		const isRoot = path === 'CHANGELOG.md' || /^viewers\/[^/]+\/CHANGELOG\.md$/u.test(path);
		const findings = checkChangelog(readFileSync(file, 'utf8'), { requireReleases: isRoot });
		for (const { line, heading, problem } of findings) {
			failures++;
			console.log(`::error file=${relative(ROOT, file)},line=${line}::${heading} ${problem}`);
		}
	}

	if (failures > 0) {
		console.log(
			`\n${failures} stub changelog section(s). Changelogs here are prepend-only ` +
				'(prune-releases.mjs culls the tags they were generated from), so these cannot be ' +
				'regenerated away: fix the release pipeline, then remove exactly the stub sections by hand.',
		);
		process.exit(1);
	}

	console.log(`OK: ${targets.length} changelog(s) checked, no stub sections.`);
}

if (process.argv[1]?.endsWith('check-changelog-sections.mjs')) {
	main();
}
