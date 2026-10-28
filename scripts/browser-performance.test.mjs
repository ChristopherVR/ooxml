import assert from 'node:assert/strict';
import { test } from 'node:test';
import { browserPerformance, pptxBrowserCommand } from './browser-performance.mjs';
import { plan } from './ci-plan.mjs';

test('strict local verification keeps local hardware and worker policy', () => {
	assert.deepEqual(browserPerformance({ CI: '1', OOXML_LOCAL_VERIFY: '1' }, 'win32', 16), {
		hosted: false,
		workers: 8,
		trace: 'on-first-retry',
	});
	assert.equal(browserPerformance({ CI: '1' }, 'linux', 4).workers, 2);
	assert.equal(browserPerformance({ CI: '1' }, 'linux', 4).hosted, true);
	assert.equal(browserPerformance({}, 'win32', 1).workers, 1);
	assert.equal(browserPerformance({}, 'win32', 16).trace, 'off');
	assert.equal(browserPerformance({ OOXML_E2E_WORKERS: '4' }, 'win32', 16).workers, 4);
	assert.throws(() => browserPerformance({ OOXML_E2E_WORKERS: 'NaN' }), /positive integer/);
});

test('local verification preserves spec selection and both pptx host variants', () => {
	const selected = plan(['e2e/pptx/viewer-basics.spec.ts']).pptx.e2e;
	const command = pptxBrowserCommand(selected);
	assert.match(command, /'viewer-basics.spec.ts'/);
	for (const project of selected.projects) {
		assert.ok(command.includes(`--project='${project}'`));
		assert.ok(command.includes(`--project='${project}-headless'`));
	}
	for (const format of ['docx', 'xlsx', 'teams']) {
		const entries = plan([`e2e/${format}/editing.spec.ts`]).viewerBrowser;
		assert.match(entries[0].local, /editing.spec.ts/);
		assert.doesNotMatch(entries[0].local, /--shard/);
	}
});
