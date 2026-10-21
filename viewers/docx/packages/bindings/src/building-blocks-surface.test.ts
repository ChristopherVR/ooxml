// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { createBuildingBlockSession } from './building-blocks';
import { createDocument } from '../../core/src/index';

afterEach(() => document.body.replaceChildren());

function words(text: string) {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text }] }];
	return model;
}

it('publishes navigation and status snapshots and keeps identical snapshots stable', () => {
	const session = createBuildingBlockSession({ documentModel: words('Count these words') });
	const host = document.createElement('div');
	document.body.append(host);
	const seen = vi.fn();
	const stop = session.subscribe(seen);
	try {
		session.canvasProps.mount(host);
		expect(session.state.statusProps).toMatchObject({ words: 3, zoom: 100 });
		expect(session.state.navigationProps).toMatchObject({ page: 1, pagesAvailable: false });
		const before = session.state;
		const calls = seen.mock.calls.length;
		session.handle.element.dispatchEvent(new Event('editor-state-change'));
		expect(session.state).toBe(before);
		expect(seen).toHaveBeenCalledTimes(calls);
		session.state.statusProps!.onZoomChange(125);
		expect(session.state.statusProps?.zoom).toBe(125);
		expect(session.state.statusProps).not.toBe(before.statusProps);
	} finally {
		stop();
		session.destroy();
	}
});

it('routes host workflow callbacks through update and reports a missing one as not enabled', () => {
	const model = words('Link target');
	const onOpenHyperlinkDialog = vi.fn();
	const session = createBuildingBlockSession({ documentModel: model });
	const host = document.createElement('div');
	document.body.append(host);
	try {
		session.canvasProps.mount(host);
		expect(session.state.toolbarProps?.canInsertLink).toBe(false);
		expect(session.state.toolbarProps?.onOpenHyperlinkDialog()).toBe(false);
		session.update({ documentModel: model, onOpenHyperlinkDialog });
		expect(session.state.toolbarProps?.canInsertLink).toBe(true);
		expect(session.state.toolbarProps?.onOpenHyperlinkDialog()).toBe(true);
		expect(onOpenHyperlinkDialog).toHaveBeenCalledOnce();
		session.update({ documentModel: model, onOpenHyperlinkDialog, canEdit: false });
		expect(session.state.toolbarProps?.canInsertLink).toBe(false);
		expect(session.state.toolbarProps?.onOpenHyperlinkDialog()).toBe(false);
	} finally {
		session.destroy();
	}
});
