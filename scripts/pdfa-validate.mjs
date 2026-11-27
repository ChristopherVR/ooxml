#!/usr/bin/env node
/**
 * Validates the Word PDF export's PDF/A output with veraPDF, the open-source reference
 * validator (https://verapdf.org). It exports the corpus in
 * `src/core/docx/pdf-export/__fixtures__/pdfa-corpus.ts` with `pdfA: true` (through its Vitest file,
 * so the sources run as the tests run them), then validates every file against each flavour.
 *
 *   VERAPDF=/path/to/verapdf node scripts/pdfa-validate.mjs [options]
 *
 *   --out <dir>          where the PDFs and reports go (a temporary folder by default)
 *   --flavours <list>    veraPDF flavours, comma-separated (default 2b,2u,2a,ua1)
 *   --fonts <files>      TrueType files to embed, separated by the path delimiter; also PDFA_FONTS.
 *                        Without either, a system sans-serif font when one is found, else the
 *                        corpus's synthetic font; `--fonts synthetic` always uses that one.
 *
 * Exits with 1 when a corpus file fails PDF/A-2b (the level the export declares), or when the
 * standard-fonts control file passes it (it must not). Other flavours are informational. It
 * never downloads veraPDF; see docs/docx-pdf-export.md for how to install it.
 */
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readdirSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { delimiter, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const REQUIRED = '2b';
const CONTROL = /-standard-fonts\.pdf$/;

const SYSTEM_FONTS = [
	'C:\\Windows\\Fonts\\arial.ttf',
	'/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
	'/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
	'/usr/share/fonts/TTF/DejaVuSans.ttf',
	'/Library/Fonts/Arial.ttf',
	'/System/Library/Fonts/Supplemental/Arial.ttf',
];

const HOW_TO_INSTALL = `veraPDF was not found. Set VERAPDF to the veraPDF command line (verapdf, or verapdf.bat on
Windows), or to the folder it is installed in. veraPDF needs Java 11 or later:

  1. Download https://software.verapdf.org/releases/verapdf-installer.zip and unzip it.
  2. Install the command line only, unattended:
       java -jar verapdf-izpack-installer-<version>.jar auto-install.xml
     with an auto-install.xml that selects the "veraPDF CLI" pack (see
     https://docs.verapdf.org/install/#automated-installation), or run the installer's GUI.
  3. VERAPDF=<install folder>/verapdf node scripts/pdfa-validate.mjs

This script never downloads veraPDF itself.`;

/** Parses the command line into options. */
export function parseArgs(argv, env = process.env) {
	const options = {
		out: undefined,
		flavours: ['2b', '2u', '2a', 'ua1'],
		fonts: [],
		synthetic: false,
	};
	for (let index = 0; index < argv.length; index++) {
		const arg = argv[index];
		const value = () => {
			const next = argv[++index];
			if (next === undefined) throw new Error(`${arg} needs a value.`);
			return next;
		};
		if (arg === '--out') options.out = value();
		else if (arg === '--flavours') options.flavours = value().split(',').filter(Boolean);
		else if (arg === '--fonts') {
			const list = value();
			if (list === 'synthetic') options.synthetic = true;
			else options.fonts = list.split(delimiter).filter(Boolean);
		} else throw new Error(`Unknown option ${arg}.`);
	}
	if (!options.fonts.length && !options.synthetic && env.PDFA_FONTS)
		options.fonts = env.PDFA_FONTS.split(delimiter).filter(Boolean);
	if (!options.flavours.includes(REQUIRED)) options.flavours.unshift(REQUIRED);
	return options;
}

/** The veraPDF executable `VERAPDF` names (a file or its install folder), or undefined. */
export function findVeraPdf(value, platform = process.platform) {
	if (!value) return undefined;
	if (existsSync(value) && statSync(value).isFile()) return value;
	const name = platform === 'win32' ? 'verapdf.bat' : 'verapdf';
	const inFolder = join(value, name);
	return existsSync(inFolder) ? inFolder : undefined;
}

/**
 * The per-file results of one veraPDF JSON report: whether each file is compliant and the
 * rules it failed (`clause-test: description (count)`).
 */
export function summariseReport(report) {
	const jobs = report?.report?.jobs ?? [];
	return jobs.map((job) => {
		const name = String(job.itemDetails?.name ?? '')
			.split(/[\\/]/)
			.at(-1);
		const results = [job.validationResult ?? []].flat();
		const failed = results.flatMap((result) =>
			(result.details?.ruleSummaries ?? [])
				.filter((rule) => rule.status !== 'passed')
				.map(
					(rule) =>
						`${rule.clause}-${rule.testNumber}: ${String(rule.description).trim()} (${rule.failedChecks})`,
				),
		);
		const compliant =
			results.length > 0 &&
			results.every((result) => result.compliant === true) &&
			!job.taskException;
		return {
			name,
			compliant,
			failed: job.taskException ? [`exception: ${job.taskException.exceptionMessage}`] : failed,
		};
	});
}

function exportCorpus(out, fonts) {
	const env = { ...process.env, PDFA_CORPUS_DIR: out, PDFA_FONTS: fonts.join(delimiter) };
	const run = spawn('bunx', ['vitest', 'run', 'docx/pdf-export/pdfa-corpus'], {
		cwd: join(ROOT, 'src/core'),
		env,
		stdio: 'inherit',
	});
	return run.status === 0;
}

/**
 * Runs a command. On Windows, commands are often batch files or shims (`verapdf.bat`, an npm
 * `bunx.cmd`) that only run through the shell, so the command line is quoted here instead.
 */
function spawn(command, args, options) {
	const run =
		process.platform === 'win32'
			? spawnSync([command, ...args].map((arg) => `"${arg}"`).join(' '), {
					...options,
					shell: true,
				})
			: spawnSync(command, args, options);
	if (run.error) throw run.error;
	return run;
}

function validate(verapdf, flavour, files) {
	const run = spawn(verapdf, ['--flavour', flavour, '--format', 'json', ...files], {
		encoding: 'utf8',
		maxBuffer: 256 * 1024 * 1024,
	});
	const start = run.stdout.indexOf('{');
	if (start < 0) throw new Error(`veraPDF printed no report:\n${run.stderr || run.stdout}`);
	return summariseReport(JSON.parse(run.stdout.slice(start)));
}

function main() {
	const options = parseArgs(process.argv.slice(2));
	const verapdf = findVeraPdf(process.env.VERAPDF);
	if (!verapdf) {
		console.error(HOW_TO_INSTALL);
		return 2;
	}
	const fonts =
		options.fonts.length || options.synthetic
			? options.fonts
			: SYSTEM_FONTS.filter((file) => existsSync(file)).slice(0, 1);
	const out = resolve(options.out ?? mkdtempSync(join(tmpdir(), 'ooxml-pdfa-')));
	mkdirSync(out, { recursive: true });
	console.log(`Exporting the corpus to ${out} (fonts: ${fonts.join(', ') || 'synthetic'})`);
	const exported = exportCorpus(out, fonts);
	const files = readdirSync(out)
		.filter((file) => file.endsWith('.pdf'))
		.sort()
		.map((file) => join(out, file));
	if (!files.length) {
		console.error('The corpus wrote no PDF.');
		return 1;
	}
	let failures = 0;
	for (const flavour of options.flavours) {
		console.log(`\nveraPDF --flavour ${flavour}`);
		for (const result of validate(verapdf, flavour, files)) {
			const control = CONTROL.test(result.name);
			console.log(
				`  ${result.compliant ? 'PASS' : 'FAIL'}  ${result.name}${control ? ' (control)' : ''}`,
			);
			for (const rule of result.failed) console.log(`        ${rule}`);
			if (flavour !== REQUIRED) continue;
			if (control ? result.compliant : !result.compliant) failures++;
		}
	}
	if (!exported)
		console.error('\nThe corpus test failed: the export report disagrees (see above).');
	if (failures) console.error(`\n${failures} file(s) did not meet PDF/A-${REQUIRED} as expected.`);
	return failures || !exported ? 1 : 0;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url))
	process.exitCode = main();
