#!/usr/bin/env node
// Runs the native Office acceptance scripts of one product and writes a machine-readable summary.
// The same entry point serves a developer's machine and the self-hosted CI runner
// (.github/workflows/office-acceptance.yml). See docs/office-acceptance.md.
//
//   node scripts/office-acceptance.mjs <word|excel|powerpoint|visio|all> [--out <dir>] [--require]
//
// Exit code: 0 when every script that ran passed (or the product was skipped because its Office
// app is not installed), 1 when a script failed, or when --require is set and an app is missing.
import { spawnSync } from 'node:child_process';
import { appendFileSync, createWriteStream, mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const core = join(root, 'src/core');

/**
 * Each step is one script. `{out}` in `args` is the step's output folder. `prepare` runs first
 * (fixture generators). `needs` are executables that must be on PATH and `apps` further COM
 * classes the step uses; the step is skipped when one is missing. Scripts report failure by
 * throwing or exiting 1 under `pwsh -File`, so none needs an adapter.
 */
export const PRODUCTS = {
	word: {
		apps: ['Word.Application'],
		steps: [
			{
				id: 'docx-smartart',
				script: join(core, 'scripts/docx-word-acceptance.ps1'),
				args: ['-OutDir', '{out}'],
			},
			{
				id: 'docx-chart',
				script: join(core, 'scripts/docx-word-chart-acceptance.ps1'),
				args: ['-OutDir', '{out}'],
				apps: ['Excel.Application'],
			},
			{
				id: 'shared-picture',
				script: join(core, 'scripts/shared-picture-native-acceptance.ps1'),
				prepare: {
					cwd: core,
					cmd: ['bun', 'scripts/shared-picture-native-fixtures.ts', '{out}'],
				},
				args: ['-InputDirectory', '{out}', '-EvidencePath', '{out}/evidence.json'],
				apps: ['Excel.Application'],
			},
		],
	},
	excel: {
		apps: ['Excel.Application'],
		steps: [
			{
				id: 'xlsx-writer',
				script: join(core, 'scripts/xlsx-excel-acceptance.ps1'),
				args: ['-OutDir', '{out}'],
			},
			{ id: 'xlsx-pivot', script: join(root, 'scripts/xlsx-pivot-acceptance.ps1'), args: [] },
			{ id: 'xlsx-slicer', script: join(root, 'scripts/xlsx-slicer-acceptance.ps1'), args: [] },
			{
				id: 'xlsx-sparkline',
				script: join(root, 'scripts/xlsx-sparkline-acceptance.ps1'),
				args: [],
			},
			{
				id: 'xlsx-print',
				script: join(root, 'scripts/xlsx-print-acceptance.ps1'),
				args: [],
				needs: ['pdfinfo', 'python'],
			},
			{
				id: 'xlsx-print-format',
				script: join(root, 'scripts/xlsx-print-format-acceptance.ps1'),
				args: [],
				needs: ['python'],
			},
		],
	},
	powerpoint: {
		apps: ['PowerPoint.Application', 'Excel.Application'],
		steps: [
			{
				id: 'pptx-classic-chart',
				script: join(core, 'scripts/pptx-classic-chart-acceptance.ps1'),
				args: ['-OutDir', '{out}'],
			},
		],
	},
	visio: {
		apps: ['Visio.Application'],
		steps: [
			{
				id: 'visio-new-drawing',
				script: join(root, 'scripts/record-visio-new-drawing.ps1'),
				prepare: {
					cwd: root,
					cmd: ['bun', 'scripts/create-visio-new-drawing-fixtures.ts', '{out}'],
				},
				args: ['-OutputDirectory', '{out}'],
			},
			{
				id: 'visio-size-position',
				script: join(root, 'scripts/record-visio-size-position.ps1'),
				prepare: {
					cwd: root,
					cmd: ['bun', 'scripts/create-visio-size-position-fixtures.ts', '{out}'],
				},
				args: ['-OutputDirectory', '{out}'],
			},
		],
	},
};

function git(...args) {
	const r = spawnSync('git', args, { cwd: root, encoding: 'utf8' });
	return r.status === 0 ? r.stdout.trim() : '';
}

let shellName;
function powershell(...args) {
	shellName ??=
		spawnSync('pwsh', ['-NoProfile', '-Command', '1'], { stdio: 'ignore' }).status === 0
			? 'pwsh'
			: 'powershell';
	return spawnSync(shellName, ['-NoProfile', '-NonInteractive', ...args], {
		cwd: root,
		encoding: 'utf8',
		maxBuffer: 256 * 1024 * 1024,
	});
}

/** The version and build of an installed Office app, or null when its COM class is absent. */
function probeApp(progId) {
	const script = `$t=[type]::GetTypeFromProgID('${progId}'); if(-not $t){exit 3}; $a=$null; try{ $a=New-Object -ComObject '${progId}'; $b=''; try{$b=[string]$a.Build}catch{}; Write-Output ($a.Version+'|'+$b) } finally { if($a){ try{$a.Quit()}catch{} } }`;
	const r = powershell('-Command', script);
	if (r.status === 3) return null;
	if (r.status !== 0)
		return { progId, version: 'unknown', build: '', error: (r.stderr || '').trim() };
	const [version = '', build = ''] = r.stdout.trim().split('|');
	return { progId, version, build };
}

function onPath(exe) {
	return (
		spawnSync(process.platform === 'win32' ? 'where' : 'which', [exe], { stdio: 'ignore' })
			.status === 0
	);
}

function notice(message) {
	console.log(process.env.GITHUB_ACTIONS ? `::notice::${message}` : message);
}

/** Counts from a script's output ("All N documents ...", "FAILED: a, b"), else per script. */
function countOf(output, ok) {
	const all = /All (\d+) (?:documents|workbooks)/u.exec(output);
	const failed = /FAILED: (.+)/u.exec(output);
	if (all) return { passed: Number(all[1]), failed: 0 };
	if (failed) return { passed: 0, failed: failed[1].split(',').length };
	return ok ? { passed: 1, failed: 0 } : { passed: 0, failed: 1 };
}

function runStep(step, dir, log) {
	const out = join(dir, step.id);
	mkdirSync(out, { recursive: true });
	const sub = (s) => s.replaceAll('{out}', out);
	const started = Date.now();
	if (step.prepare) {
		const [cmd, ...rest] = step.prepare.cmd.map(sub);
		const p = spawnSync(cmd, rest, {
			cwd: step.prepare.cwd,
			encoding: 'utf8',
			shell: process.platform === 'win32',
		});
		const text = `${p.stdout ?? ''}${p.stderr ?? ''}`;
		log.write(text);
		if (p.status !== 0) {
			return {
				status: 'failed',
				exitCode: p.status ?? -1,
				detail: 'fixture generation failed',
				durationMs: Date.now() - started,
				passed: 0,
				failed: 1,
			};
		}
	}
	const p = powershell('-File', step.script, ...step.args.map(sub));
	const text = `${p.stdout ?? ''}${p.stderr ?? ''}`;
	log.write(text);
	const ok = p.status === 0;
	return {
		status: ok ? 'passed' : 'failed',
		exitCode: p.status ?? -1,
		durationMs: Date.now() - started,
		...countOf(text, ok),
	};
}

const relative = (path) => path.slice(root.length + 1).replaceAll('\\', '/');

export function runProduct(name, outRoot, { require = false } = {}) {
	const product = PRODUCTS[name];
	const dir = join(outRoot, name);
	mkdirSync(dir, { recursive: true });
	const summary = {
		product: name,
		date: new Date().toISOString(),
		commit: git('rev-parse', 'HEAD'),
		runner: process.env.RUNNER_NAME ?? process.env.COMPUTERNAME ?? '',
		office: [],
		status: 'passed',
		totals: { passed: 0, failed: 0, skipped: 0 },
		scripts: [],
	};
	const finish = () => {
		writeFileSync(join(dir, 'summary.json'), `${JSON.stringify(summary, null, '\t')}\n`);
		if (process.env.GITHUB_STEP_SUMMARY) {
			const lines = [
				`### Office acceptance: ${name} (${summary.status})`,
				'',
				'| script | status | passed | failed |',
				'| --- | --- | --- | --- |',
				...summary.scripts.map((s) => `| ${s.id} | ${s.status} | ${s.passed} | ${s.failed} |`),
			];
			appendFileSync(process.env.GITHUB_STEP_SUMMARY, `${lines.join('\n')}\n\n`);
		}
		return summary;
	};
	const present = new Map();
	for (const progId of product.apps) {
		const info = process.platform === 'win32' ? probeApp(progId) : null;
		if (!info) {
			summary.status = require ? 'failed' : 'skipped';
			summary.reason = `${progId} is not installed on this machine`;
			summary.totals.skipped = product.steps.length;
			notice(`Skipping ${name}: ${summary.reason}.`);
			return finish();
		}
		present.set(progId, info);
		summary.office.push(info);
	}
	const log = createWriteStream(join(dir, 'run.log'));
	for (const step of product.steps) {
		const missing = [
			...(step.needs ?? []).filter((exe) => !onPath(exe)),
			...(step.apps ?? []).filter((id) => !present.has(id) && !probeApp(id)),
		];
		if (missing.length > 0) {
			const detail = `needs ${missing.join(', ')}`;
			summary.scripts.push({
				id: step.id,
				script: relative(step.script),
				status: 'skipped',
				detail,
				passed: 0,
				failed: 0,
			});
			summary.totals.skipped++;
			notice(`Skipping ${step.id}: ${detail}.`);
			continue;
		}
		console.log(`== ${name}/${step.id}`);
		log.write(`== ${step.id}\n`);
		const r = runStep(step, dir, log);
		summary.scripts.push({ id: step.id, script: relative(step.script), ...r });
		summary.totals.passed += r.passed;
		summary.totals.failed += r.failed;
		console.log(`   ${r.status} (exit ${r.exitCode})`);
		if (r.status === 'failed') summary.status = 'failed';
	}
	log.end();
	return finish();
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
	const argv = process.argv.slice(2);
	const outIndex = argv.indexOf('--out');
	const outRoot = resolve(
		outIndex >= 0 ? (argv[outIndex + 1] ?? '') : join(root, 'office-acceptance-results'),
	);
	const require = argv.includes('--require');
	const target = argv.find((a, i) => !a.startsWith('--') && argv[i - 1] !== '--out');
	const names = target === 'all' ? Object.keys(PRODUCTS) : [target];
	if (!target || names.some((n) => !(n in PRODUCTS))) {
		console.error(
			`Usage: node scripts/office-acceptance.mjs <${Object.keys(PRODUCTS).join('|')}|all> [--out <dir>] [--require]`,
		);
		process.exit(2);
	}
	mkdirSync(outRoot, { recursive: true });
	const results = names.map((n) => runProduct(n, outRoot, { require }));
	writeFileSync(
		join(outRoot, 'summary.json'),
		`${JSON.stringify({ date: new Date().toISOString(), commit: git('rev-parse', 'HEAD'), products: results }, null, '\t')}\n`,
	);
	process.exit(results.some((r) => r.status === 'failed') ? 1 : 0);
}
