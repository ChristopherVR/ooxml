#!/usr/bin/env node
/**
 * workflow-lint.mjs: the supply-chain rules of the workflows that actionlint and zizmor do not check.
 *
 * 1. `bunx`, `bun x`, `npx` and `npm exec` run only a package some workspace declares, so bun.lock
 *    holds its integrity. A spec with a version (`tool@1.2.3`) or an undeclared package fails.
 * 2. A job whose token can write (a `write` scope other than `id-token`, `write-all`, or no
 *    `permissions` at all) runs no third-party code: its steps use only actions of `actions/` and
 *    `github/`, local composite actions made of such steps, and shell that starts no package
 *    manager or installed tool. A repository script it runs with `node` may import only `node:`
 *    builtins and relative files, and may not name a third-party command.
 * 3. A step that runs third-party code receives no secret other than the job's `GITHUB_TOKEN`.
 *
 *   node scripts/workflow-lint.mjs
 */

import { readFileSync, readdirSync } from 'node:fs';
import { builtinModules } from 'node:module';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { parse } from 'yaml';

import { parseLock } from './check-lockfile-drift.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const FIRST_PARTY_OWNERS = new Set(['actions', 'github']);
const BUILTINS = new Set(builtinModules);
// Commands that run code installed from a registry, or a registry runner itself.
const THIRD_PARTY_COMMANDS = [
	'bun',
	'bunx',
	'npx',
	'pnpm',
	'yarn',
	'pip',
	'pip3',
	'pipx',
	'cosign',
];
const THIRD_PARTY_SHELL = [
	new RegExp(`(?:^|[\\s;&|(\`])(${THIRD_PARTY_COMMANDS.join('|')})(?=\\s|$)`, 'mu'),
	/\bnpm\s+(?:install|i|ci|exec|x|run|run-script|rebuild|test|start|restart)\b/u,
	/node_modules\//u,
	/\bpython3?\s+-m\s+pip\b/u,
	/\b(?:curl|wget)\b[^\n|]*\|\s*(?:ba|z)?sh\b/u,
];
const RUNNER =
	/(?:^|[\s;&|(`])(bunx|npx|bun\s+x|npm\s+exec|pnpm\s+dlx|yarn\s+dlx)\s+([^\n;&|)]*)/gmu;

/** The commands a workspace can run from its own node_modules: declared packages and their bins. */
export function lockedCommands(lockText) {
	const lock = parseLock(lockText);
	const commands = new Set();
	for (const workspace of Object.values(lock.workspaces ?? {})) {
		for (const field of ['dependencies', 'devDependencies', 'optionalDependencies']) {
			for (const name of Object.keys(workspace[field] ?? {})) {
				commands.add(name);
				const bin = lock.packages?.[name]?.[2]?.bin;
				if (bin && typeof bin === 'object') for (const key of Object.keys(bin)) commands.add(key);
			}
		}
	}
	return commands;
}

/** Each `bunx`-style invocation in a shell script: the runner and the package spec it names. */
export function runnerInvocations(script) {
	const found = [];
	for (const [, runner, rest] of script.matchAll(RUNNER)) {
		const words = rest.trim().split(/\s+/u);
		let spec;
		for (let index = 0; index < words.length && !spec; index++) {
			const word = words[index];
			if (word === '-p' || word === '--package') spec = words[index + 1];
			else if (word.startsWith('--package=')) spec = word.slice('--package='.length);
			else if (!word.startsWith('-')) spec = word;
		}
		if (spec)
			found.push({ runner: runner.replace(/\s+/u, ' '), spec: spec.replace(/^['"]|['"]$/gu, '') });
	}
	return found;
}

/** Why a package spec is not locked, or null when bun.lock covers it. */
export function unlockedReason(spec, locked) {
	if (spec.includes('$')) return `\`${spec}\` is computed, so bun.lock cannot cover it`;
	const at = spec.indexOf('@', spec.startsWith('@') ? 1 : 0);
	if (at > 0)
		return `\`${spec}\` names a version outside bun.lock; declare it and drop the version`;
	return locked.has(spec) ? null : `\`${spec}\` is not a dependency of any workspace in bun.lock`;
}

/** The scopes a job's token can write, from its own `permissions` or the workflow's. */
export function writeScopes(permissions) {
	if (permissions === undefined) return ['the default token (no permissions block)'];
	if (permissions === null || permissions === 'read-all') return [];
	if (typeof permissions === 'string') return permissions === 'write-all' ? ['write-all'] : [];
	return Object.entries(permissions)
		.filter(([scope, level]) => level === 'write' && scope !== 'id-token')
		.map(([scope]) => scope);
}

/** Bare module specifiers and third-party command names in JavaScript source. */
function sourceProblems(source) {
	const problems = [];
	const specifiers = /(?:\bfrom\s*|\bimport\s*\(\s*|\brequire\s*\(\s*)['"]([^'"]+)['"]/gu;
	for (const [, specifier] of source.matchAll(specifiers)) {
		if (specifier.startsWith('.') || specifier.startsWith('node:') || BUILTINS.has(specifier))
			continue;
		problems.push(`imports \`${specifier}\``);
	}
	const names = [...THIRD_PARTY_COMMANDS, 'git-cliff', 'pkg-pr-new'].join('|');
	for (const [, name] of source.matchAll(new RegExp(`['"\`](${names})['"\`]`, 'gu'))) {
		problems.push(`runs \`${name}\``);
	}
	if (source.includes('node_modules/.bin')) problems.push('runs a tool from node_modules');
	return problems;
}

/** Why a repository script (and every relative module it imports) runs third-party code. */
function scriptReason(file, root, seen = new Set()) {
	if (seen.has(file)) return null;
	seen.add(file);
	let source;
	try {
		source = readFileSync(file, 'utf8');
	} catch {
		return `${relative(root, file)} is not in the repository`;
	}
	const [problem] = sourceProblems(source);
	if (problem) return `${relative(root, file).replaceAll('\\', '/')} ${problem}`;
	if (file.endsWith('.json')) return null;
	for (const [, specifier] of source.matchAll(/\bfrom\s*['"](\.[^'"]+)['"]/gu)) {
		const reason = scriptReason(resolve(dirname(file), specifier), root, seen);
		if (reason) return reason;
	}
	return null;
}

/** Why a `run:` script runs third-party code, or null. */
function shellReason(script, root) {
	const pattern = THIRD_PARTY_SHELL.find((candidate) => candidate.test(script));
	if (pattern) return `runs \`${script.match(pattern)[0].trim()}\``;
	const [inline] = sourceProblems(script.replace(/\bnode\s+\S+\.m?js\b/gu, ''));
	if (inline) return `inline script ${inline}`;
	for (const [, path] of script.matchAll(/\bnode\s+(?:--\S+\s+)*([\w./-]+\.m?js)\b/gu)) {
		const reason = scriptReason(resolve(root, path), root);
		if (reason) return reason;
	}
	return null;
}

/** Why a step runs third-party code, or null when it is first-party throughout. */
export function stepReason(step, root, readAction = defaultReadAction) {
	if (step.uses) {
		if (step.uses.startsWith('./')) {
			const action = readAction(root, step.uses);
			if (action?.runs?.using !== 'composite') return `local action ${step.uses} is not composite`;
			for (const inner of action.runs.steps ?? []) {
				const reason = stepReason(inner, root, readAction);
				if (reason) return `${step.uses} ${reason}`;
			}
			return null;
		}
		const owner = step.uses.split('/')[0];
		return FIRST_PARTY_OWNERS.has(owner) ? null : `uses ${step.uses.split('@')[0]}`;
	}
	return typeof step.run === 'string' ? shellReason(step.run, root) : null;
}

function defaultReadAction(root, uses) {
	for (const name of ['action.yml', 'action.yaml']) {
		try {
			return parse(readFileSync(join(root, uses, name), 'utf8'));
		} catch {
			// try the next name
		}
	}
	return null;
}

const secretsIn = (value) =>
	[...JSON.stringify(value ?? {}).matchAll(/secrets\.([A-Za-z0-9_]+)/gu)]
		.map(([, name]) => name)
		.filter((name) => name !== 'GITHUB_TOKEN');

/** Findings for one parsed workflow. */
export function lintWorkflow(workflow, { file, root, locked, readAction }) {
	const findings = [];
	for (const [jobId, job] of Object.entries(workflow?.jobs ?? {})) {
		const at = (step, index) =>
			`${file} job ${jobId} step "${step.name ?? step.uses ?? index + 1}"`;
		const scopes = writeScopes(job.permissions ?? workflow.permissions);
		if (job.uses && !job.uses.startsWith('./') && scopes.length > 0) {
			findings.push(
				`${file} job ${jobId} calls ${job.uses} with a token that writes ${scopes.join(', ')}`,
			);
		}
		const inherited = secretsIn([workflow.env, job.env]);
		(job.steps ?? []).forEach((step, index) => {
			for (const { runner, spec } of runnerInvocations(step.run ?? '')) {
				const reason = unlockedReason(spec, locked);
				if (reason) findings.push(`${at(step, index)} runs \`${runner} ${spec}\`: ${reason}`);
			}
			const reason = stepReason(step, root, readAction);
			if (!reason) return;
			if (scopes.length > 0) {
				findings.push(
					`${at(step, index)} ${reason} in a job whose token writes ${scopes.join(', ')}`,
				);
			}
			const secrets = [...inherited, ...secretsIn([step.env, step.with])];
			if (secrets.length > 0) {
				findings.push(`${at(step, index)} ${reason} and receives ${secrets.join(', ')}`);
			}
		});
	}
	return findings;
}

/** Findings for the composite actions: only the runner rule applies. */
export function lintAction(action, { file, locked }) {
	return (action?.runs?.steps ?? []).flatMap((step, index) =>
		runnerInvocations(step.run ?? '').flatMap(({ runner, spec }) => {
			const reason = unlockedReason(spec, locked);
			return reason ? [`${file} step ${index + 1} runs \`${runner} ${spec}\`: ${reason}`] : [];
		}),
	);
}

/** Every finding in the repository at `root`. */
export function lintRepository(root = ROOT) {
	const locked = lockedCommands(readFileSync(join(root, 'bun.lock'), 'utf8'));
	const findings = [];
	const workflows = join(root, '.github/workflows');
	for (const name of readdirSync(workflows)
		.filter((entry) => /\.ya?ml$/u.test(entry))
		.sort()) {
		const file = `.github/workflows/${name}`;
		const workflow = parse(readFileSync(join(workflows, name), 'utf8'));
		findings.push(...lintWorkflow(workflow, { file, root, locked }));
	}
	const actions = join(root, '.github/actions');
	for (const name of readdirSync(actions).sort()) {
		const uses = `./.github/actions/${name}`;
		const action = defaultReadAction(root, uses);
		findings.push(...lintAction(action, { file: `.github/actions/${name}`, locked }));
	}
	return findings;
}

function main() {
	const findings = lintRepository();
	for (const finding of findings) {
		console.error(process.env.GITHUB_ACTIONS ? `::error::${finding}` : finding);
	}
	if (findings.length > 0) process.exit(1);
	console.log('Workflow policy: every runner is locked and no third-party step can write.');
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
