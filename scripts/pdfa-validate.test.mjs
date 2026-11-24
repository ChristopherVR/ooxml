import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { delimiter, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { findVeraPdf, parseArgs, summariseReport } from './pdfa-validate.mjs';

const SCRIPT = fileURLToPath(new URL('./pdfa-validate.mjs', import.meta.url));

test('parseArgs always validates PDF/A-2b and reads fonts from the environment', () => {
	assert.deepEqual(parseArgs(['--flavours', '2u,ua1'], {}).flavours, ['2b', '2u', 'ua1']);
	assert.deepEqual(parseArgs([], { PDFA_FONTS: `a.ttf${delimiter}b.ttf` }).fonts, [
		'a.ttf',
		'b.ttf',
	]);
	assert.equal(parseArgs(['--fonts', 'synthetic'], { PDFA_FONTS: 'a.ttf' }).synthetic, true);
	assert.deepEqual(parseArgs(['--fonts', 'synthetic'], { PDFA_FONTS: 'a.ttf' }).fonts, []);
	assert.throws(() => parseArgs(['--nope'], {}), /Unknown option/);
});

test('findVeraPdf accepts the command or its install folder', (t) => {
	const folder = mkdtempSync(join(tmpdir(), 'ooxml-verapdf-'));
	t.after(() => rmSync(folder, { recursive: true, force: true }));
	writeFileSync(join(folder, 'verapdf'), '');
	assert.equal(findVeraPdf(undefined), undefined);
	assert.equal(findVeraPdf(folder, 'linux'), join(folder, 'verapdf'));
	assert.equal(findVeraPdf(join(folder, 'verapdf'), 'linux'), join(folder, 'verapdf'));
	assert.equal(findVeraPdf(folder, 'win32'), undefined);
});

test('summariseReport lists each file with the rules it failed', () => {
	const report = {
		report: {
			jobs: [
				{
					itemDetails: { name: 'C:\\out\\text.pdf' },
					validationResult: [{ compliant: true, details: { ruleSummaries: [] } }],
				},
				{
					itemDetails: { name: '/out/chart.pdf' },
					validationResult: [
						{
							compliant: false,
							details: {
								ruleSummaries: [
									{
										status: 'failed',
										clause: '6.2.11.4.1',
										testNumber: 1,
										description: 'Fonts embedded ',
										failedChecks: 2,
									},
								],
							},
						},
					],
				},
				{ itemDetails: { name: 'x.pdf' }, taskException: { exceptionMessage: 'broken' } },
			],
		},
	};
	assert.deepEqual(summariseReport(report), [
		{ name: 'text.pdf', compliant: true, failed: [] },
		{ name: 'chart.pdf', compliant: false, failed: ['6.2.11.4.1-1: Fonts embedded (2)'] },
		{ name: 'x.pdf', compliant: false, failed: ['exception: broken'] },
	]);
});

test('without VERAPDF it explains how to install veraPDF and downloads nothing', () => {
	const run = spawnSync(process.execPath, [SCRIPT], {
		encoding: 'utf8',
		env: { ...process.env, VERAPDF: '' },
	});
	assert.equal(run.status, 2);
	assert.match(run.stderr, /veraPDF was not found/);
	assert.match(run.stderr, /never downloads/);
});
