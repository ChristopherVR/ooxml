// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest';

import { stageKeepsElementIds, stripElementIdMarkers } from './stage-element-ids';

function buildStage(): HTMLElement {
	const stage = document.createElement('div');
	stage.innerHTML = [
		'<div data-element-id="a"><span>A</span></div>',
		'<div data-element-id="b" data-pptx-element="true">B</div>',
		'<div class="plain">C</div>',
	].join('');
	return stage;
}

describe('stageKeepsElementIds', () => {
	it('keeps the markers on the editing canvas, the show stage and Morph layers only', () => {
		expect(stageKeepsElementIds({ interactive: true })).toBeTruthy();
		expect(stageKeepsElementIds({ presenting: true })).toBeTruthy();
		expect(stageKeepsElementIds({ preserveElementIds: true })).toBeTruthy();
		expect(stageKeepsElementIds({})).toBeFalsy();
		expect(
			stageKeepsElementIds({ interactive: false, presenting: false, preserveElementIds: false }),
		).toBeFalsy();
	});
});

describe('stripElementIdMarkers', () => {
	it('removes every data-element-id under the stage and reports the count', () => {
		const stage = buildStage();
		expect(stripElementIdMarkers(stage)).toBe(2);
		expect(stage.querySelectorAll('[data-element-id]')).toHaveLength(0);
		// Content and unrelated attributes are untouched.
		expect(stage.textContent).toBe('ABC');
		expect(stage.querySelector('[data-pptx-element]')).not.toBeNull();
	});

	it('also clears the root node', () => {
		const node = document.createElement('div');
		node.setAttribute('data-element-id', 'root');
		node.innerHTML = '<span data-element-id="child"></span>';
		expect(stripElementIdMarkers(node)).toBe(2);
		expect(node.hasAttribute('data-element-id')).toBeFalsy();
	});

	it('is a no-op on a stage without markers', () => {
		const stage = document.createElement('div');
		stage.innerHTML = '<div class="plain">X</div>';
		expect(stripElementIdMarkers(stage)).toBe(0);
		expect(stage.textContent).toBe('X');
	});
});
