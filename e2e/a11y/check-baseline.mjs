#!/usr/bin/env node
/**
 * Keeps the accessibility baselines honest: they may only shrink.
 *
 *   node check-baseline.mjs --base <ref>    fail if baseline/axe.json allows more than at <ref>,
 *                                           expected-failures.json lists a gap <ref> did not, or
 *                                           samples-expected.json excuses a notice <ref> did not
 *                                           or one that names no open item of the parity checklist
 *   node check-baseline.mjs --update [--from <dir>]
 *                                           lower baseline/axe.json to what a run found (its
 *                                           axe.json files); never raises an entry. The counts are
 *                                           CI's: off Linux, pass --from with the unpacked
 *                                           a11y-report artifact of a CI run
 *   node check-baseline.mjs --update --add  also record violations the baseline does not list yet:
 *                                           only for a new surface or rule, and the --base check
 *                                           still refuses it against the base branch
 */
import { execFileSync } from 'node:child_process';
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const AXE = join(HERE, 'baseline', 'axe.json');
const GAPS = join(HERE, 'expected-failures.json');
const NOTICES = join(HERE, 'samples-expected.json');
const CHECKLIST = join(HERE, '..', '..', 'docs', 'parity-checklist.md');
const read = (path) => JSON.parse(readFileSync(path, 'utf8'));

const git = (...args) =>
	execFileSync('git', args, { cwd: HERE, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });

/** The commit `ref` names. A ref git cannot resolve is an error, never an empty base. */
function commitOf(ref) {
	try {
		return git('rev-parse', '--verify', '--quiet', `${ref}^{commit}`).trim();
	} catch {
		throw new Error(`git cannot find the base ${ref}: fetch it or pass another --base.`);
	}
}

/** The file at `commit`, or undefined when it does not exist there yet. */
function atCommit(commit, path) {
	const repo = relative(git('rev-parse', '--show-toplevel').trim(), path).replaceAll('\\', '/');
	try {
		git('cat-file', '-e', `${commit}:${repo}`);
	} catch {
		console.log(`${repo} is new since the base: nothing to compare it with.`);
		return undefined;
	}
	return JSON.parse(git('show', `${commit}:${repo}`));
}

/** Every way `now` allows more than `before`: a higher count, or a surface or rule it lacked. */
export function grown(before, now) {
	const problems = [];
	for (const [key, rules] of Object.entries(now.allowed ?? {}))
		for (const [rule, count] of Object.entries(rules)) {
			const was = before.allowed?.[key]?.[rule];
			if (was === undefined) problems.push(`${key}: ${rule} is new (${count} nodes)`);
			else if (count > was) problems.push(`${key}: ${rule} rose from ${was} to ${count} nodes`);
		}
	return problems;
}

/** Rules the baseline allows without naming the work package that removes them. */
export const unowned = (baseline) =>
	[...new Set(Object.values(baseline.allowed ?? {}).flatMap(Object.keys))].filter(
		(rule) => !baseline.owners?.[rule],
	);

/**
 * Every way `now` excuses more than `before`: a known gap the base did not list, one that does not
 * say which failure it expects, or one that expects a different failure than at the base (which
 * could match an unrelated one).
 */
export function loosenedGaps(before, now) {
	const problems = [];
	for (const [key, gap] of Object.entries(now.gaps ?? {})) {
		const was = before.gaps?.[key];
		if (!was) problems.push(`${key} is new`);
		else if (!gap.fails) problems.push(`${key} does not say which failure it expects`);
		else if (was.fails !== undefined && gap.fails !== was.fails)
			problems.push(`${key} expects a different failure`);
	}
	return problems;
}

const oneLine = (text) => text.replace(/\s+/gu, ' ').trim();

/** The open items of the parity checklist (`- [ ] ...`), each as one line of text. */
export function openGaps(markdown) {
	return markdown
		.split(/^(?=- \[[ x]\] )/mu)
		.filter((item) => item.startsWith('- [ ] '))
		.map(oneLine);
}

/**
 * Every load notice a sample is allowed (samples-expected.json) that does not quote an open item
 * of the parity checklist: a notice is excused only while the gap behind it is on the list.
 */
export function unmappedNotices(expected, markdown) {
	const gaps = openGaps(markdown);
	const problems = [];
	for (const [sample, entries] of Object.entries(expected.notices ?? {}))
		for (const entry of entries) {
			if (!entry?.notice) problems.push(`${sample}: an entry has no notice`);
			else if (!entry.gap || !gaps.some((gap) => gap.includes(oneLine(entry.gap))))
				problems.push(
					`${sample}: "${entry.notice}" quotes no open item of docs/parity-checklist.md`,
				);
		}
	return problems;
}

/**
 * Every way `now` excuses more load notices than `before`: a standing note it newly ignores, or a
 * notice a sample was not allowed at the base. A sample the base did not list is new and brings
 * its own notices (each must still quote an open gap).
 */
export function grownNotices(before, now) {
	const problems = [];
	const ignored = new Set(before.informational ?? []);
	for (const start of now.informational ?? [])
		if (!ignored.has(start)) problems.push(`informational "${start}" is new`);
	for (const [sample, entries] of Object.entries(now.notices ?? {})) {
		const was = before.notices?.[sample];
		if (!was) continue;
		const known = new Set(was.map((entry) => entry.notice));
		for (const entry of entries)
			if (!known.has(entry.notice)) problems.push(`${sample}: "${entry.notice}" is new`);
	}
	return problems;
}

/** The baseline lowered to `observed` ({ key: { rule: nodes } }), raised only with `add`. */
export function lowered(baseline, observed, add = false) {
	const allowed = structuredClone(baseline.allowed ?? {});
	for (const [key, rules] of Object.entries(observed)) {
		const old = allowed[key];
		if (!old && !add) continue;
		const next = {};
		for (const [rule, count] of Object.entries(rules)) {
			const was = old?.[rule];
			if (was !== undefined) next[rule] = Math.min(was, count);
			else if (add) next[rule] = count;
		}
		if (Object.keys(next).length) allowed[key] = sortKeys(next);
		else delete allowed[key];
	}
	return { ...baseline, allowed: sortKeys(allowed) };
}

const sortKeys = (object) =>
	Object.fromEntries(Object.entries(object).sort(([a], [b]) => a.localeCompare(b)));

function observations(dir) {
	const found = {};
	if (!existsSync(dir)) return found;
	for (const entry of readdirSync(dir, { withFileTypes: true, recursive: true })) {
		if (!entry.isFile() || entry.name !== 'axe.json') continue;
		const { key, violations } = read(join(entry.parentPath, entry.name));
		found[key] = Object.fromEntries(violations.map((v) => [v.rule, v.nodes]));
	}
	return found;
}

function main(argv) {
	const value = (flag) => (argv.includes(flag) ? argv[argv.indexOf(flag) + 1] : undefined);
	if (argv.includes('--update')) {
		const from = value('--from');
		// Fonts differ by platform and move some counts, so the baseline follows CI's runner.
		if (!from && process.platform !== 'linux')
			throw new Error('Lower the baseline from a CI run: --update --from <unpacked a11y-report>.');
		const observed = observations(from ?? join(HERE, 'test-results'));
		if (!Object.keys(observed).length)
			throw new Error('No axe results found: run axe.spec.ts first.');
		const next = lowered(read(AXE), observed, argv.includes('--add'));
		writeFileSync(AXE, `${JSON.stringify(next, null, '\t')}\n`);
		console.log(`baseline/axe.json updated from ${Object.keys(observed).length} surfaces.`);
		return 0;
	}
	const base = value('--base');
	if (!base) throw new Error('Pass --base <ref> or --update.');
	const commit = commitOf(base);
	const problems = unowned(read(AXE)).map((rule) => `baseline/axe.json: ${rule} has no owner`);
	const axeBefore = atCommit(commit, AXE);
	if (axeBefore)
		problems.push(...grown(axeBefore, read(AXE)).map((p) => `baseline/axe.json: ${p}`));
	const gapsBefore = atCommit(commit, GAPS);
	if (gapsBefore)
		problems.push(
			...loosenedGaps(gapsBefore, read(GAPS)).map((p) => `expected-failures.json: ${p}`),
		);
	const notices = read(NOTICES);
	problems.push(
		...unmappedNotices(notices, readFileSync(CHECKLIST, 'utf8')).map(
			(p) => `samples-expected.json: ${p}`,
		),
	);
	const noticesBefore = atCommit(commit, NOTICES);
	if (noticesBefore)
		problems.push(
			...grownNotices(noticesBefore, notices).map((p) => `samples-expected.json: ${p}`),
		);
	if (problems.length) {
		console.error(
			`The accessibility baselines may only shrink (compared with ${base}):\n  ${problems.join('\n  ')}\nFix the regression instead of allowing it.`,
		);
		return 1;
	}
	console.log(`Accessibility baselines are no larger than at ${base}.`);
	return 0;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1]))
	process.exitCode = main(process.argv.slice(2));
