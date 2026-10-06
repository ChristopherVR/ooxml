import { describe, expect, it } from 'vitest';

import {
	EMPTY_SUITE_STATE,
	activateSuiteTab,
	closeSuiteTab,
	newSuiteTabId,
	openSuiteTab,
	parseSuiteState,
	setSuiteTabFramework,
	suiteTabTitles,
} from './tab-state.js';

function three() {
	let state = openSuiteTab(EMPTY_SUITE_STATE, 'word', 'react', 't1');
	state = openSuiteTab(state, 'word', 'react', 't2');
	return openSuiteTab(state, 'excel', 'vue', 't3');
}

describe('suite tab state', () => {
	it('activates a newly opened tab', () => {
		expect(three().active).toBe('t3');
		expect(three().tabs).toHaveLength(3);
	});

	it('numbers the second tab of the same app on', () => {
		const titles = suiteTabTitles(three().tabs, (id) => id);
		expect([...titles]).toEqual([
			['t1', 'word'],
			['t2', 'word 2'],
			['t3', 'excel'],
		]);
	});

	it('closes the active tab onto its right neighbour, then the left, then Home', () => {
		let state = closeSuiteTab(activateSuiteTab(three(), 't2'), 't2');
		expect(state.active).toBe('t3');
		state = closeSuiteTab(state, 't3');
		expect(state.active).toBe('t1');
		state = closeSuiteTab(state, 't1');
		expect(state).toEqual(EMPTY_SUITE_STATE);
	});

	it('keeps the active tab when a background tab closes', () => {
		const state = closeSuiteTab(activateSuiteTab(three(), 't1'), 't3');
		expect(state.active).toBe('t1');
		expect(state.tabs.map((t) => t.id)).toEqual(['t1', 't2']);
	});

	it('falls back to Home for an unknown id', () => {
		expect(activateSuiteTab(three(), 'nope').active).toBeNull();
	});

	it('changes the framework of the named tab only', () => {
		const state = setSuiteTabFramework(three(), 't2', 'vue');
		expect(state.tabs.map((t) => t.framework)).toEqual(['react', 'vue', 'vue']);
	});

	it('drops malformed and duplicate tabs and a stale active id when parsing', () => {
		const state = parseSuiteState({
			tabs: [
				{ id: 'a', app: 'word', framework: 'react' },
				{ id: 'a', app: 'x', framework: 'y' },
				5,
				null,
			],
			active: 'gone',
		});
		expect(state).toEqual({ tabs: [{ id: 'a', app: 'word', framework: 'react' }], active: null });
		expect(parseSuiteState('junk')).toEqual(EMPTY_SUITE_STATE);
		expect(parseSuiteState({ tabs: 'no' })).toEqual(EMPTY_SUITE_STATE);
	});

	it('never reuses a live tab id', () => {
		expect(newSuiteTabId(closeSuiteTab(three(), 't1').tabs)).toBe('t4');
	});
});
