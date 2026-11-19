// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { dispatchEditorKey } from './editor-key-dispatch';
import { mapEditorKey } from './editor-keymap';
import { handleNotesFieldKey } from './notes/notes-keys';
import {
	announceSlideObject,
	attachSlideKeyboardFocus,
	editSlideObjectText,
	slideKeyFocus,
} from './slide-keyboard-focus';

let detach: () => void = () => {};

function mountViewer(): {
	viewport: HTMLElement;
	ribbonTab: HTMLElement;
	notes: HTMLTextAreaElement;
	slideRow: HTMLElement;
} {
	document.body.innerHTML = `
		<button id="before">before</button>
		<div data-pptx-editor-chrome tabindex="0">
			<div data-pptx-chrome="ribbon">
				<button role="tab" aria-selected="false">File</button>
				<button role="tab" aria-selected="true" id="home">Home</button>
			</div>
			<aside data-pptx-chrome="slides"><button id="row" aria-current="true">1</button></aside>
			<div data-pptx-viewport>
				<div aria-roledescription="slide" aria-label="Slide 1">
					<div data-element-id="a" aria-label="Title"></div>
					<div data-element-id="b" aria-label="Picture of a boat"></div>
				</div>
			</div>
			<div data-pptx-chrome="notes"><textarea id="notes"></textarea></div>
		</div>`;
	const viewport = document.querySelector<HTMLElement>('[data-pptx-viewport]')!;
	detach = attachSlideKeyboardFocus(viewport, () => true);
	return {
		viewport,
		ribbonTab: document.getElementById('home')!,
		notes: document.querySelector('textarea')!,
		slideRow: document.getElementById('row')!,
	};
}

const press = (target: Element, key: string, init: KeyboardEventInit = {}): KeyboardEvent => {
	const event = new KeyboardEvent('keydown', {
		key,
		bubbles: true,
		composed: true,
		cancelable: true,
		...init,
	});
	target.dispatchEvent(event);
	return event;
};

beforeEach(() => {
	// jsdom lays nothing out; every pane counts as showing.
	vi.spyOn(HTMLElement.prototype, 'getClientRects').mockReturnValue([
		{} as DOMRect,
	] as unknown as DOMRectList);
});

afterEach(() => {
	detach();
	detach = () => {};
	vi.restoreAllMocks();
	document.body.innerHTML = '';
});

describe('Tab and Enter in the keymap', () => {
	it('leaves Tab to the browser everywhere except inside the slide', () => {
		expect(mapEditorKey({ key: 'Tab' }).action).toBeNull();
		expect(mapEditorKey({ key: 'Tab' }, { slideFocus: 'outside' }).action).toBeNull();
		expect(mapEditorKey({ key: 'Tab' }, { slideFocus: 'pane' }).action).toBeNull();
		expect(mapEditorKey({ key: 'Tab', shiftKey: true }, { slideFocus: 'pane' }).action).toBeNull();
	});

	it('walks the objects once inside the slide', () => {
		expect(mapEditorKey({ key: 'Tab' }, { slideFocus: 'objects' }).action).toBe(
			'cycleSelectionNext',
		);
		expect(mapEditorKey({ key: 'Tab', shiftKey: true }, { slideFocus: 'objects' }).action).toBe(
			'cycleSelectionPrev',
		);
	});

	it('goes in with Enter only from the slide pane', () => {
		expect(mapEditorKey({ key: 'Enter' }, { slideFocus: 'pane' }).action).toBe(
			'cycleSelectionNext',
		);
		expect(mapEditorKey({ key: 'Enter' }, { slideFocus: 'outside' }).action).toBeNull();
		expect(mapEditorKey({ key: 'Enter' }, { slideFocus: 'objects' }).action).toBeNull();
		expect(
			mapEditorKey({ key: 'Enter' }, { slideFocus: 'pane', canEdit: false }).action,
		).toBeNull();
	});
});

describe('going into and out of the slide', () => {
	const dispatch = (
		target: Element,
		key: string,
		hasSelection = false,
		handlers: Parameters<typeof dispatchEditorKey>[2] = {},
	) => {
		const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true });
		Object.defineProperty(event, 'target', { value: target });
		return { event, action: dispatchEditorKey(event, { hasSelection }, handlers) };
	};

	it('is one labelled Tab stop until the user goes in', () => {
		const { viewport } = mountViewer();
		expect(viewport.tabIndex).toBe(0);
		expect(viewport.getAttribute('role')).toBe('group');
		expect(viewport.getAttribute('aria-label')).toBe('Slide 1');
		expect(viewport.getAttribute('aria-keyshortcuts')).toContain('F6');
		expect(slideKeyFocus(viewport)).toBe('pane');
		const { event, action } = dispatch(viewport, 'Tab');
		expect(action).toBeNull();
		expect(event.defaultPrevented).toBe(false);
	});

	it('Enter selects the first object, Tab then walks, Escape comes back out', () => {
		const { viewport } = mountViewer();
		const cycleSelectionNext = vi.fn();
		dispatch(viewport, 'Enter', false, { cycleSelectionNext });
		expect(cycleSelectionNext).toHaveBeenCalledOnce();
		expect(slideKeyFocus(viewport)).toBe('objects');

		const tab = dispatch(viewport, 'Tab', true, { cycleSelectionNext });
		expect(tab.event.defaultPrevented).toBe(true);
		expect(cycleSelectionNext).toHaveBeenCalledTimes(2);

		// The first Escape clears the selection; the second leaves the slide.
		dispatch(viewport, 'Escape', true);
		expect(slideKeyFocus(viewport)).toBe('objects');
		dispatch(viewport, 'Escape', false);
		expect(slideKeyFocus(viewport)).toBe('pane');
		expect(dispatch(viewport, 'Tab').event.defaultPrevented).toBe(false);
	});

	it('leaves Enter and Tab to a focused link or button on the slide', () => {
		const { viewport } = mountViewer();
		const link = viewport.appendChild(document.createElement('a'));
		expect(slideKeyFocus(link)).toBe('outside');
		const enter = dispatch(link, 'Enter');
		expect(enter.action).toBeNull();
		expect(enter.event.defaultPrevented).toBe(false);
	});

	it('finds the viewport inside a scrolling wrapper, also when it mounts later', async () => {
		document.body.innerHTML = '<div data-pptx-editor-chrome><main id="main"></main></div>';
		const main = document.getElementById('main')!;
		detach = attachSlideKeyboardFocus(main, () => true);
		const viewport = document.createElement('div');
		viewport.setAttribute('data-pptx-viewport', '');
		main.append(viewport);
		await Promise.resolve();
		expect(viewport.tabIndex).toBe(0);
		expect(viewport.getAttribute('role')).toBe('group');
	});

	it('does not go into a slide with nothing on it', () => {
		const { viewport } = mountViewer();
		viewport.querySelectorAll('[data-element-id]').forEach((node) => node.remove());
		dispatch(viewport, 'Enter');
		expect(slideKeyFocus(viewport)).toBe('pane');
	});

	it('goes in on a pointer press and focuses the slide', () => {
		const { viewport } = mountViewer();
		viewport.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
		expect(slideKeyFocus(viewport)).toBe('objects');
		expect(document.activeElement).toBe(viewport);
	});

	it('leaves the objects when focus moves elsewhere', () => {
		const { viewport, notes } = mountViewer();
		viewport.setAttribute('data-pptx-slide-objects', '');
		viewport.focus();
		notes.focus();
		expect(slideKeyFocus(viewport)).toBe('pane');
	});

	it('announces the object Tab reached', () => {
		const { viewport } = mountViewer();
		announceSlideObject(document, 'b');
		const status = viewport.closest('[data-pptx-editor-chrome]')!.querySelector('[role="status"]');
		expect(status?.textContent).toBe('Picture of a boat');
	});
});

describe('F6 between panes', () => {
	it('visits ribbon, slide, notes and slide list clockwise, and back with Shift', () => {
		const { viewport, ribbonTab, notes, slideRow } = mountViewer();
		viewport.focus();
		press(viewport, 'F6');
		expect(document.activeElement).toBe(notes);
		press(notes, 'F6');
		expect(document.activeElement).toBe(slideRow);
		press(slideRow, 'F6');
		expect(document.activeElement).toBe(ribbonTab);
		press(ribbonTab, 'F6');
		expect(document.activeElement).toBe(viewport);
		expect(slideKeyFocus(viewport)).toBe('objects');
		press(viewport, 'F6', { shiftKey: true });
		expect(document.activeElement).toBe(ribbonTab);
		expect(slideKeyFocus(viewport)).toBe('pane');
	});

	it('claims F6 only for a viewer that holds focus, and never over a dialog', () => {
		const { ribbonTab } = mountViewer();
		const before = document.getElementById('before')!;
		before.focus();
		expect(press(before, 'F6').defaultPrevented).toBe(false);
		expect(document.activeElement).toBe(before);

		const dialog = document.createElement('div');
		dialog.setAttribute('aria-modal', 'true');
		const field = dialog.appendChild(document.createElement('input'));
		ribbonTab.closest('[data-pptx-editor-chrome]')!.append(dialog);
		field.focus();
		expect(press(field, 'F6').defaultPrevented).toBe(false);
		expect(document.activeElement).toBe(field);
	});
});

describe('the notes field', () => {
	it('lets Tab and F6 through, and Escape goes back to the slide', () => {
		const { viewport, notes, slideRow } = mountViewer();
		const keys = (event: Event) => handleNotesFieldKey(event as KeyboardEvent);
		notes.addEventListener('keydown', keys);
		notes.focus();
		const tab = press(notes, 'Tab');
		expect(tab.defaultPrevented).toBe(false);
		press(notes, 'F6');
		expect(document.activeElement).toBe(slideRow);
		notes.focus();
		press(notes, 'Escape');
		expect(document.activeElement).toBe(viewport);
		expect(slideKeyFocus(viewport)).toBe('pane');
	});

	it('commits on Ctrl+Enter and goes back to the slide, but leaves an IME alone', () => {
		const { viewport, notes } = mountViewer();
		const commit = vi.fn();
		notes.addEventListener('keydown', (event) => handleNotesFieldKey(event, commit));
		notes.focus();
		press(notes, 'Escape', { isComposing: true });
		expect(document.activeElement).toBe(notes);
		const enter = press(notes, 'Enter', { ctrlKey: true });
		expect(commit).toHaveBeenCalledTimes(1);
		expect(enter.defaultPrevented).toBe(true);
		expect(document.activeElement).toBe(viewport);
	});
});

describe('Enter on a selected object', () => {
	it('double-clicks the object on the slide, as the mouse does to edit its text', () => {
		const { viewport } = mountViewer();
		const shape = viewport.querySelector('[data-element-id="b"]')!;
		const onDoubleClick = vi.fn();
		shape.addEventListener('dblclick', onDoubleClick);
		editSlideObjectText(document, 'b');
		editSlideObjectText(document, null);
		editSlideObjectText(document, 'missing');
		expect(onDoubleClick).toHaveBeenCalledTimes(1);
	});
});

describe('keys inside a dialog', () => {
	it('never reach the editor keymap, unless the dialog holds the viewer', () => {
		const { viewport } = mountViewer();
		const dialog = document.createElement('div');
		dialog.setAttribute('role', 'dialog');
		dialog.setAttribute('aria-modal', 'true');
		const button = dialog.appendChild(document.createElement('button'));
		document.body.append(dialog);
		const remove = vi.fn();
		const pressDelete = (target: Element) => {
			target.addEventListener(
				'keydown',
				(event) => dispatchEditorKey(event, { hasSelection: true }, { delete: remove }),
				{ once: true },
			);
			press(target, 'Delete');
		};
		pressDelete(button);
		expect(remove).not.toHaveBeenCalled();

		const host = document.createElement('dialog');
		document.body.append(host);
		host.append(viewport.closest('[data-pptx-editor-chrome]')!);
		pressDelete(viewport);
		expect(remove).toHaveBeenCalledTimes(1);
	});
});
