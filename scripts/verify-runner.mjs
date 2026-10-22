/**
 * The runner behind scripts/verify.mjs: a small dependency-aware parallel scheduler and the build
 * stamps that let verify skip a build whose inputs have not changed.
 *
 * A step is `{ id, cmd, cwd, needs?: string[], label?, skipIf?(): string | false, onSuccess?() }`.
 * A step starts when every step it needs has finished successfully and a slot is free. After a
 * failure nothing new starts (running steps are left to finish, so their output is not cut off
 * mid-write and no process tree has to be killed); steps that depended on the failure are skipped.
 */
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

/**
 * Run `steps`. `exec(step)` resolves `{ code, output }`. `onStart` / `onDone` / `onTick` are
 * optional reporting hooks. Resolves `{ results, failed }`, where `results` has one entry per step
 * in the order they finished: `{ step, status: 'ok' | 'failed' | 'skipped' | 'blocked', ms, output,
 * reason }`.
 */
export async function runSteps(steps, { jobs = 1, exec, onStart, onDone } = {}) {
	const ids = new Set(steps.map((step) => step.id));
	for (const step of steps)
		for (const need of step.needs ?? [])
			if (!ids.has(need)) throw new Error(`step ${step.id} needs unknown step ${need}`);
	const pending = new Set(steps);
	const state = new Map(); // id -> 'ok' | 'failed' | 'skipped' | 'blocked'
	const results = [];
	const running = new Set();
	let failed = false;

	const finish = (step, status, extra = {}) => {
		state.set(step.id, status);
		const result = { step, status, ms: 0, output: '', reason: '', ...extra };
		results.push(result);
		onDone?.(result);
	};

	return new Promise((resolve, reject) => {
		const pump = () => {
			let progressed = true;
			while (progressed) {
				progressed = false;
				for (const step of [...pending]) {
					const needs = step.needs ?? [];
					if (needs.some((id) => state.get(id) === 'failed' || state.get(id) === 'blocked')) {
						pending.delete(step);
						finish(step, 'blocked', { reason: 'a step it needs failed' });
						progressed = true;
						continue;
					}
					if (failed || running.size >= jobs) continue;
					if (!needs.every((id) => state.has(id))) continue;
					pending.delete(step);
					progressed = true;
					const reason = step.skipIf?.();
					if (reason) {
						finish(step, 'skipped', { reason });
						continue;
					}
					running.add(step);
					onStart?.(step, running);
					const at = Date.now();
					Promise.resolve(exec(step)).then(
						({ code, output }) => {
							running.delete(step);
							const ok = code === 0;
							if (ok) step.onSuccess?.();
							else failed = true;
							finish(step, ok ? 'ok' : 'failed', { ms: Date.now() - at, output });
							pump();
						},
						(error) => {
							running.delete(step);
							failed = true;
							finish(step, 'failed', { ms: Date.now() - at, output: String(error) });
							pump();
						},
					);
				}
			}
			if (running.size === 0) {
				// Whatever is left was never started because of a failure.
				for (const step of pending) finish(step, 'blocked', { reason: 'an earlier step failed' });
				pending.clear();
				resolve({ results, failed });
			}
		};
		try {
			pump();
		} catch (error) {
			reject(error);
		}
	});
}

/** Default number of concurrent heavy steps: vitest and tsc are multi-threaded themselves. */
export const defaultJobs = (parallelism) => Math.max(1, Math.floor(parallelism / 4));

const IGNORED_INPUT =
	/(?:\.test\.[cm]?[jt]sx?$|\.spec\.[cm]?[jt]sx?$|\/__tests__\/|\/__fixtures__\/|\.mdx?$|\/docs\/|\/README)/;

/**
 * A hash of a build's inputs: the tracked files under `dirs` (path, mode and blob id from
 * `git ls-files -s`, which is what a clean tree holds; untracked files are not seen), minus tests,
 * fixtures and docs, plus the contents of `files` (lockfile, shared configs) and any `extra`
 * strings (for example the stamp of the build this one consumes). `lsFiles(dir)` returns the
 * `ls-files -s` text of a directory.
 */
export function hashInputs({ dirs, files = [], extra = [], lsFiles, root }) {
	const hash = createHash('sha256');
	for (const dir of dirs) {
		const lines = lsFiles(dir)
			.split('\n')
			.filter((line) => line && !IGNORED_INPUT.test(line.split('\t')[1] ?? ''));
		hash.update(`dir ${dir}\n${lines.join('\n')}\n`);
	}
	for (const file of files) {
		const path = join(root, file);
		hash.update(`file ${file}\n`);
		if (existsSync(path)) hash.update(readFileSync(path));
	}
	for (const value of extra) hash.update(`extra ${value}\n`);
	return hash.digest('hex');
}

export const STAMP = '.verify-build-stamp';

/** Why `dist` can be reused, or false. */
export function stampMatches(distDir, hash) {
	const stamp = join(distDir, STAMP);
	if (!existsSync(distDir) || !existsSync(stamp)) return false;
	return readFileSync(stamp, 'utf8').trim() === hash;
}

/** Drop the stamp before a build starts so a half-written `dist` is never trusted. */
export function clearStamp(distDir) {
	rmSync(join(distDir, STAMP), { force: true });
}

export function writeStamp(distDir, hash) {
	mkdirSync(distDir, { recursive: true });
	writeFileSync(join(distDir, STAMP), `${hash}\n`);
}
