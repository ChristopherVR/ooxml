// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { mountViewer } from './binding';
import { demoDocument } from 'ooxml-core/visio/ui';

afterEach(() => document.body.replaceChildren());

function setup() {
	const host = document.createElement('div');
	document.body.append(host);
	const root = mountViewer(host, { document: demoDocument }).element.shadowRoot!;
	const canvas = root.querySelector<HTMLElement>('.viewport')!;
	const deepActive = (): Element | null => {
		let active = root.activeElement;
		while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
		return active;
	};
	/** The region the keyboard is in now, across shadow roots and slots. */
	const inRegion = (selector: string): boolean => {
		const region = root.querySelector(selector);
		for (let node: Node | null = deepActive(); node;)
			if (node === region) return true;
			else
				node =
					(node as Element).assignedSlot ?? node.parentNode ?? (node as ShadowRoot).host ?? null;
		return false;
	};
	const press = (shiftKey = false) =>
		(deepActive() ?? canvas).dispatchEvent(
			new KeyboardEvent('keydown', { key: 'F6', shiftKey, bubbles: true, composed: true }),
		);
	return { canvas, root, inRegion, press };
}

describe('Visio region keys', () => {
	it('walks F6 from the drawing window round the window and back', () => {
		const { canvas, inRegion, press } = setup();
		canvas.focus();
		press();
		expect(inRegion('.page-bar')).toBe(true);
		press();
		expect(inRegion('office-ui-status-bar')).toBe(true);
		press();
		expect(inRegion('office-ui-ribbon')).toBe(true);
		press();
		expect(inRegion('#shapes-pane')).toBe(true);
		press();
		expect(inRegion('.viewport')).toBe(true);
	});

	it('walks Shift+F6 the other way', () => {
		const { canvas, inRegion, press } = setup();
		canvas.focus();
		press(true);
		expect(inRegion('#shapes-pane')).toBe(true);
		press(true);
		expect(inRegion('office-ui-ribbon')).toBe(true);
		press(true);
		expect(inRegion('office-ui-status-bar')).toBe(true);
		press(true);
		expect(inRegion('.page-bar')).toBe(true);
		press(true);
		expect(inRegion('.viewport')).toBe(true);
	});

	it('passes over the task pane while it is closed', () => {
		const { canvas, root, inRegion, press } = setup();
		expect(root.querySelector<HTMLElement>('#inspector-pane')!.hidden).toBe(true);
		canvas.focus();
		press();
		expect(inRegion('#inspector-pane')).toBe(false);
	});
});
