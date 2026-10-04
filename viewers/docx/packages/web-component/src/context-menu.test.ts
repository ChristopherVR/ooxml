// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createContextMenu } from './context-menu';
import type { RibbonAction } from './ribbon-action';
import { schema } from './schema';

afterEach(() => document.body.replaceChildren());

function setup() {
	const frame = document.createElement('div');
	const canvas = document.createElement('div');
	document.body.append(frame, canvas);
	const doc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('Hello')])]);
	const state = EditorState.create({ doc });
	const view = new EditorView(document.createElement('div'), {
		state: state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 4))),
	});
	const run = vi.fn<(action: RibbonAction) => void>();
	const focusDocument = vi.fn();
	const menu = createContextMenu({
		container: frame,
		target: canvas,
		view: () => view,
		focusDocument,
		readOnly: () => false,
		locale: () => 'en',
		run,
		warn: vi.fn(),
	});
	return { frame, menu, run, focusDocument };
}

describe('context menu', () => {
	it('shows the shared menu element with the translated items', async () => {
		const { frame, menu } = setup();
		menu.open(10, 20);
		const element = frame.querySelector('office-ui-context-menu');
		expect(element).not.toBeNull();
		expect(menu.isOpen).toBe(true);
		expect(menu.element).toBe(element);
		await (element as unknown as { updateComplete: Promise<unknown> }).updateComplete;
		const labels = [...(element?.shadowRoot?.querySelectorAll('[role="menuitem"]') ?? [])].map(
			(row) => row.textContent,
		);
		expect(labels).toEqual(expect.arrayContaining(['Cut', 'Copy', 'Paste']));
	});

	it('runs the chosen command and closes', () => {
		const { frame, menu, run, focusDocument } = setup();
		menu.open(10, 20);
		frame
			.querySelector('office-ui-context-menu')
			?.dispatchEvent(new CustomEvent('office-menu-request', { detail: { id: 'link-insert' } }));
		expect(run).toHaveBeenCalledWith({ type: 'link' });
		expect(menu.isOpen).toBe(false);
		expect(frame.querySelector('office-ui-context-menu')).toBeNull();
		expect(focusDocument).toHaveBeenCalled();
	});

	it('closes on dismissal, keeping focus where an outside click put it', () => {
		const { frame, menu, focusDocument } = setup();
		menu.open(10, 20);
		frame
			.querySelector('office-ui-context-menu')
			?.dispatchEvent(new CustomEvent('office-menu-close', { detail: { reason: 'outside' } }));
		expect(menu.isOpen).toBe(false);
		expect(focusDocument).not.toHaveBeenCalled();
	});
});
