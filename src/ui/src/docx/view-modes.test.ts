// @vitest-environment jsdom
import { createDocument } from 'ooxml-core/docx';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, expect, it, vi } from 'vitest';
import type { PageControllerHost } from './page-controller';
import { ViewModes } from './view-modes';
import { mountView as mountRead } from './view-read';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

function surface() {
	const element = document.createElement('div');
	const frame = document.createElement('div');
	frame.className = 'dve-frame';
	const canvas = document.createElement('div');
	canvas.tabIndex = 0;
	canvas.dataset.documentView = 'draft';
	const paper = document.createElement('div');
	paper.className = 'dve-paper';
	canvas.append(paper);
	frame.append(canvas);
	element.append(frame);
	document.body.append(element);
	const model = createDocument();
	const host = {
		element,
		paper: () => paper,
		toolbar: () => undefined,
		statusBar: () => undefined,
		view: () => undefined,
		model: () => model,
		setModel: vi.fn(),
		locale: () => 'en',
		canEditOutsideBody: () => true,
		edited: vi.fn(),
		reportError: vi.fn(),
		printLayout: () => undefined,
		printEditing: vi.fn(),
		beforeViewChange: vi.fn(),
		configureRibbon: vi.fn(),
		refreshEditable: vi.fn(),
		refreshControls: vi.fn(),
	} satisfies PageControllerHost;
	return { host, frame, canvas, paper, modes: new ViewModes(host) };
}

it('disposes and mounts before updating ribbon, editability and controls', async () => {
	const { host, canvas, frame, modes } = surface();
	const calls: string[] = [];
	host.beforeViewChange.mockImplementation(() =>
		calls.push(`before ${canvas.dataset.documentView}`),
	);
	host.printEditing.mockImplementation((active: boolean) => calls.push(`print ${active}`));
	host.configureRibbon.mockImplementation(() =>
		calls.push(`ribbon ${frame.querySelectorAll('style').length}`),
	);
	host.refreshEditable.mockImplementation(() => calls.push(`editable ${modes.mode}`));
	host.refreshControls.mockImplementation(() => calls.push(`controls ${modes.active}`));
	modes.select('web');
	await vi.waitFor(() => expect(modes.mode).toBe('web'));
	modes.select('print');
	expect(calls).toEqual([
		'before draft',
		'print false',
		'ribbon 1',
		'editable web',
		'controls web',
		'before web',
		'print true',
		'ribbon 0',
		'editable print',
		'controls print',
	]);
});

it('discards older asynchronous selections after a newer synchronous choice', async () => {
	const { host, frame, modes } = surface();
	modes.select('web');
	modes.select('read');
	modes.select('print');
	await Promise.all([import('./view-web'), import('./view-read')]);
	expect(modes.active).toBe('print');
	expect(frame.querySelectorAll('style')).toHaveLength(0);
	expect(host.refreshControls).toHaveBeenCalledTimes(1);
	expect(host.reportError).not.toHaveBeenCalled();
});

it('exits Focus over Read before restoring the layout that preceded Read', async () => {
	const { frame, modes } = surface();
	modes.select('print');
	modes.select('read');
	await vi.waitFor(() => expect(modes.active).toBe('read'));
	const controls = frame.querySelector('.dve-reader-controls');
	expect(controls).not.toBeNull();
	modes.select('focus');
	await vi.waitFor(() => expect(modes.active).toBe('focus'));
	expect(modes.escape()).toBe(true);
	expect(modes.active).toBe('read');
	expect(frame.querySelector('.dve-reader-controls')).toBe(controls);
	expect(frame.hasAttribute('data-focus-view')).toBe(false);
	expect(modes.escape()).toBe(true);
	expect(modes.active).toBe('print');
	expect(frame.querySelector('.dve-reader-controls')).toBeNull();
	expect(modes.escape()).toBe(false);
});

it('invalidates a view load on disconnect and permits selecting it after reconnect', async () => {
	const { host, frame, modes } = surface();
	modes.select('read');
	host.element.remove();
	modes.disconnect();
	document.body.append(host.element);
	await import('./view-read');
	expect(modes.active).toBe('draft');
	expect(frame.querySelector('.dve-reader-controls')).toBeNull();
	expect(host.refreshControls).not.toHaveBeenCalled();
	modes.select('read');
	await vi.waitFor(() => expect(modes.active).toBe('read'));
	modes.disconnect();
	expect(frame.querySelector('.dve-reader-controls')).toBeNull();
});

it('pages by one spread and gap, tolerates rounding and clamps after resize', async () => {
	const { host, frame, canvas, paper } = surface();
	const schema = new Schema({
		nodes: {
			doc: { content: 'paragraph+' },
			paragraph: { content: 'text*', toDOM: () => ['p', 0] },
			text: {},
		},
	});
	const view = new EditorView(
		{ mount: paper.appendChild(document.createElement('div')) },
		{
			state: EditorState.create({ schema }),
		},
	);
	let width = 600;
	let flowWidth = 1865;
	view.dom.style.columnGap = '32px';
	Object.defineProperty(view.dom, 'scrollWidth', { get: () => flowWidth });
	Object.defineProperty(paper, 'clientWidth', { get: () => width });
	let resize = () => {};
	const disconnect = vi.fn();
	vi.stubGlobal(
		'ResizeObserver',
		class {
			constructor(callback: () => void) {
				resize = callback;
			}
			observe() {}
			disconnect = disconnect;
		},
	);
	const presentation = mountRead({
		...host,
		paper,
		canvas,
		frame,
		toolbar: undefined,
		leave: vi.fn(),
	});
	try {
		presentation.sync?.(view);
		const status = frame.querySelector('[role="status"]');
		const buttons = frame.querySelectorAll<HTMLButtonElement>('.dve-reader-controls button');
		const [previous, next] = buttons;
		if (!previous || !next) throw new Error('Expected reader paging buttons');
		expect(status?.textContent).toBe('Page 1 of 3');
		expect(previous.getAttribute('aria-disabled')).toBe('true');
		await new Promise(requestAnimationFrame);
		next.click();
		expect(status?.textContent).toBe('Page 2 of 3');
		expect(paper.scrollLeft).toBe(632);
		next.focus();
		const press = (key: string) =>
			next.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
		expect(press('PageUp')).toBe(false);
		expect(paper.scrollLeft).toBe(0);
		expect(press(' ')).toBe(true);
		expect(press('Enter')).toBe(true);
		next.click();
		next.click();
		expect(paper.scrollLeft).toBe(1264);
		expect(next.getAttribute('aria-disabled')).toBe('true');
		// A shorter viewport can add columns even when the paper width is unchanged.
		flowWidth = 2497;
		resize();
		await new Promise(requestAnimationFrame);
		expect(status?.textContent).toBe('Page 3 of 4');
		expect(next.getAttribute('aria-disabled')).toBe('false');
		width = flowWidth = 500;
		resize();
		await new Promise(requestAnimationFrame);
		expect(paper.scrollLeft).toBe(0);
		expect(status?.textContent).toBe('Page 1 of 1');
		expect(next.getAttribute('aria-disabled')).toBe('true');
	} finally {
		presentation.dispose();
		view.destroy();
	}
	expect(disconnect).toHaveBeenCalledOnce();
});
