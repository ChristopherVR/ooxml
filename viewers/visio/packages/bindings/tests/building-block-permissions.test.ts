import { afterEach, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { createBuildingBlockSession, effectiveMode } from '../src/building-blocks';
import { toolbarGroups } from '../src/toolbar-controls';
import { DiagramCanvas, Toolbar, useViewerBuildingBlocks } from '../../react/src/building-blocks';
import { createViewerBuildingBlocks } from '../src/vanilla-building-blocks';

afterEach(() => {
	document.body.replaceChildren();
});

it('is read-only from creation until the host grants permission', () => {
	const session = createBuildingBlockSession();
	expect(session.controller.canEdit).toBe(false);
	session.update({ canEdit: true });
	expect(session.controller.canEdit).toBe(true);
	session.update({});
	expect(session.controller.canEdit).toBe(false);
	session.destroy();
	expect(createBuildingBlockSession({ canEdit: true }).controller.canEdit).toBe(true);
});

it('reports view mode for a drawing that cannot be saved back, whatever the host permits', () => {
	const session = createBuildingBlockSession({ canEdit: true });
	// Permission is granted but nothing editable is open (a legacy file has no source either).
	expect(session.controller.state.canEdit).toBe(true);
	expect(effectiveMode(session.controller.state)).toBe('view');
	expect(session.toolbar().mode).toBe('view');
	expect(session.toolbar().canDraw).toBe(false);
	session.destroy();
});

it('offers host dialog buttons only for the callbacks the host supplied', () => {
	const session = createBuildingBlockSession();
	const open = vi.fn();
	const none = session.toolbar({});
	expect(none.canOpenSettings || none.canOpenShareDialog || none.canOpenPageSetup).toBe(false);
	expect(toolbarGroups(none).some((group) => group.id === 'file')).toBe(false);
	const given = session.toolbar({ onOpenSettings: open, onOpenShareDialog: open });
	const file = toolbarGroups(given).find((group) => group.id === 'file')!;
	expect(file.controls.map((control) => control.id)).toEqual(['share', 'settings']);
	given.onOpenSettings?.();
	expect(open).toHaveBeenCalledOnce();
	session.destroy();
});

it('keeps the toolbar groups optional and enabled-state aware', () => {
	const session = createBuildingBlockSession();
	const props = session.toolbar();
	const groups = toolbarGroups(props, ['history', 'zoom']);
	expect(groups.map((group) => group.id)).toEqual(['history', 'zoom']);
	const history = groups[0]!.controls;
	expect(history.map((control) => control.id)).toEqual(['undo', 'redo']);
	expect(history.every((control) => control.kind === 'button' && control.disabled)).toBe(true);
	session.destroy();
});

it('applies host permission before the first React render and mounts the library toolbar', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const seen: boolean[] = [];
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	function Host() {
		blocks = useViewerBuildingBlocks({ canEdit: false, onOpenSettings: () => {} });
		seen.push(blocks.inspectorProps.controller.canEdit);
		return createElement(
			'div',
			null,
			createElement(DiagramCanvas, blocks.canvasProps),
			createElement(Toolbar, blocks.toolbarProps),
		);
	}
	try {
		await act(async () => root.render(createElement(Host)));
		expect(seen.every((value) => value === false)).toBe(true);
		const toolbar = host.querySelector('[role="toolbar"]')!;
		for (const id of ['undo', 'redo', 'zoom-in', 'zoom-out', 'fit', 'delete', 'copy', 'settings'])
			expect(toolbar.querySelector(`[data-control="${id}"]`), id).not.toBeNull();
		expect(toolbar.querySelector<HTMLButtonElement>('[data-control="delete"]')!.disabled).toBe(
			true,
		);
		const markClean = vi.spyOn(blocks.inspectorProps.controller, 'markClean');
		blocks.markClean();
		expect(markClean).toHaveBeenCalledOnce();
	} finally {
		await act(async () => root.unmount());
	}
});

it('gives vanilla the same blocks: mount helpers, shared toolbar, optional pieces and disposal', () => {
	const blocks = createViewerBuildingBlocks({ canEdit: true, onOpenFile: () => {} });
	const surface = document.createElement('div');
	const toolbar = document.createElement('div');
	const pages = document.createElement('select');
	const inspector = document.createElement('div');
	const shapes = document.createElement('div');
	document.body.append(surface, toolbar, pages, inspector, shapes);
	try {
		expect(blocks.controller.canEdit).toBe(true);
		expect(blocks.mode).toBe('view');
		blocks.mountCanvas(surface);
		const stopToolbar = blocks.mountToolbar(toolbar, ['history', 'file']);
		blocks.mountPageNavigation(pages);
		blocks.mountInspector(inspector);
		blocks.mountShapes(shapes);
		expect(surface.shadowRoot!.querySelector('[part="canvas"]')).not.toBeNull();
		expect(shapes.shadowRoot!.querySelector('.shapes-pane')).not.toBeNull();
		expect(toolbar.getAttribute('role')).toBe('toolbar');
		expect(
			[...toolbar.querySelectorAll('[data-control]')].map((e) => e.getAttribute('data-control')),
		).toEqual(['open', 'undo', 'redo']);
		const undo = toolbar.querySelector<HTMLButtonElement>('[data-control="undo"]')!;
		// A state change refreshes in place: the same element survives, so focus does too.
		blocks.controller.setZoom(2);
		expect(toolbar.querySelector('[data-control="undo"]')).toBe(undo);
		expect(blocks.toolbarProps.zoom).toBe(2);
		stopToolbar();
		expect(toolbar.childElementCount).toBe(0);
		blocks.update({ canEdit: false });
		expect(blocks.controller.canEdit).toBe(false);
	} finally {
		blocks.destroy();
	}
	expect(surface.shadowRoot!.children).toHaveLength(0);
	expect(shapes.shadowRoot!.children).toHaveLength(0);
	expect(() => blocks.controller.setZoom(3)).toThrow();
});
