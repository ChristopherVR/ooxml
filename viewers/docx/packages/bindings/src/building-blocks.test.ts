// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { createBuildingBlockSession } from './building-blocks';
import { createDocument, saveDocx } from '../../core/src/index';
afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});
it('keeps the engine on repeated refs, guards retained commands and cleans subscriptions', async () => {
	const session = createBuildingBlockSession({ documentModel: createDocument(), canEdit: true });
	const host = document.createElement('div');
	document.body.append(host);
	const snapshots = vi.fn();
	const stop = session.subscribe(snapshots);
	try {
		session.canvasProps.mount(host);
		const element = session.handle.element;
		expect(session.state.toolbarProps?.canEdit).toBe(true);
		session.canvasProps.mount(host);
		expect(session.handle.element).toBe(element);
		session.state.toolbarProps!.onBold();
		expect(session.state.toolbarProps?.bold).toBe(true);
		const retained = session.state.toolbarProps!;
		session.update({ canEdit: false });
		expect(session.state.mode).toBe('view');
		expect(retained.onItalic()).toBe(false);
		session.canvasProps.mount(null);
		expect(retained.onItalic()).toBe(false);
		expect(host.childElementCount).toBe(0);
		stop();
		const before = snapshots.mock.calls.length;
		element.dispatchEvent(new Event('editor-state-change'));
		expect(snapshots).toHaveBeenCalledTimes(before);
		expect(() => session.handle.element).toThrow('not mounted');
	} finally {
		session.destroy();
	}
});
it('reloads identical bytes after canvas replacement and recovers a rejected load', async () => {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Reattached document' }] }];
	const bytes = await saveDocx(model);
	const session = createBuildingBlockSession({ bytes });
	const first = document.createElement('div');
	const second = document.createElement('div');
	document.body.append(first, second);
	try {
		session.canvasProps.mount(first);
		await vi.waitFor(() =>
			expect(session.handle.element.documentModel?.blocks[0]).toMatchObject({
				runs: [{ text: 'Reattached document' }],
			}),
		);
		const initial = session.handle.element;
		await expect(session.handle.load(new Uint8Array([0x50, 0x4b, 3, 4]))).rejects.toThrow(/.+/);
		expect(session.state.error).toBeInstanceOf(Error);
		session.canvasProps.mount(second);
		await vi.waitFor(() =>
			expect(session.handle.element.documentModel?.blocks[0]).toMatchObject({
				runs: [{ text: 'Reattached document' }],
			}),
		);
		expect(session.handle.element).not.toBe(initial);
		expect(first.childElementCount).toBe(0);
		expect(session.state.error).toBeNull();
		expect(session.state.loading).toBe(false);
	} finally {
		session.destroy();
	}
});

it('is read-only until the host passes canEdit: true', () => {
	const session = createBuildingBlockSession({ documentModel: createDocument() });
	const host = document.createElement('div');
	document.body.append(host);
	try {
		expect(session.state.mode).toBe('view');
		session.canvasProps.mount(host);
		expect(session.state.mode).toBe('view');
		expect(session.state.toolbarProps?.canEdit).toBe(false);
		session.update({ canEdit: true });
		expect(session.state.mode).toBe('edit');
		expect(session.state.toolbarProps?.canEdit).toBe(true);
	} finally {
		session.destroy();
	}
});
