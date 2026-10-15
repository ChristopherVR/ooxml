import assert from 'node:assert/strict';
import { test } from 'node:test';
import { shellLayout } from '../apps/office-suite/src/shell-layout.js';

const base = { width: 1366, height: 768, view: 'editor', documents: 1, teamsOpen: false };

test('desktop keeps every bar', () => {
	assert.deepEqual(shellLayout(base), {
		narrow: false,
		hideTabs: false,
		collapsible: false,
		collapsed: false,
	});
});

test('a phone portrait editor folds the actions and hides a lone tab', () => {
	const layout = shellLayout({ ...base, width: 375, height: 812 });
	assert.equal(layout.narrow, true);
	assert.equal(layout.hideTabs, true);
	assert.equal(layout.collapsed, false);
});

test('two documents or Teams keep the tab strip', () => {
	assert.equal(shellLayout({ ...base, width: 375, documents: 2 }).hideTabs, false);
	assert.equal(shellLayout({ ...base, width: 375, teamsOpen: true }).hideTabs, false);
});

test('the home view never hides tabs or collapses', () => {
	const layout = shellLayout({ ...base, width: 375, height: 375, view: 'home' });
	assert.equal(layout.hideTabs, false);
	assert.equal(layout.collapsible, false);
});

test('phone landscape collapses the chrome until it is expanded', () => {
	const landscape = { ...base, width: 812, height: 375 };
	assert.equal(shellLayout(landscape).collapsed, true);
	assert.equal(shellLayout({ ...landscape, expanded: true }).collapsed, false);
	assert.equal(shellLayout({ ...landscape, expanded: true }).collapsible, true);
});

test('the breakpoints are inclusive', () => {
	assert.equal(shellLayout({ ...base, width: 640 }).narrow, true);
	assert.equal(shellLayout({ ...base, width: 641 }).narrow, false);
	assert.equal(shellLayout({ ...base, height: 500 }).collapsible, true);
	assert.equal(shellLayout({ ...base, height: 501 }).collapsible, false);
});
