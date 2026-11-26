/**
 * Wires the keyboard shortcut registry, the shortcut help dialog and the context menu to the
 * editor. Everything dispatches through the existing controllers (`routeRibbonAction`, the chrome's
 * file commands, the search panel); this module holds only the focus and layer bookkeeping.
 */
import type { RibbonElement } from './ribbon-tab-api';
import type { EditorCore } from './editor-core';
import { emit } from './events';
import { attachCanvasClick } from './canvas-click';
import { focusView } from './focus-view';
import { createContextMenu, type ContextMenu } from './context-menu';
import {
	createShortcutRegistry,
	installKeyboard,
	isMacPlatform,
	type ShortcutActions,
} from './keyboard';
import { routeRibbonAction } from './ribbon-router';
import { createShortcutHelp, shortcutRows } from './shortcut-help';

type Region = 'ribbon' | 'document' | 'status';
const REGIONS: Region[] = ['ribbon', 'document', 'status'];

/** Something Escape can close: a dialog, panel or menu. */
interface Layer {
	element: HTMLElement | undefined;
	isOpen(): boolean;
	close(): void;
	/** The layer already handles its own Escape key while focus is inside it. */
	selfHandled: boolean;
}

export function attachEditorInteractions(core: EditorCore, frame: HTMLElement): void {
	const { element, shell } = core;
	const activeElement = () => element.shadowRoot?.activeElement ?? null;
	const presentation = () =>
		core.pages.views.active === 'read' || core.pages.views.active === 'focus';
	// A read-only ProseMirror surface cannot take focus, so the canvas stands in for it.
	shell.canvas!.tabIndex = -1;
	attachCanvasClick(core);
	const toDocument = () => {
		const view = core.targetView();
		if (view?.editable) focusView(view);
		else shell.canvas?.focus();
	};

	const help = createShortcutHelp(toDocument);
	frame.append(help.element);
	const menu: ContextMenu = createContextMenu({
		container: frame,
		target: shell.canvas!,
		view: () => core.targetView(),
		focusDocument: toDocument,
		readOnly: () => !core.canEditBody(),
		locale: () => core.locale,
		run: (action) => routeRibbonAction(core, action),
		warn: (message) => emit(element, 'document-warning', message),
		customization: () => core.viewOptions.customization,
	});
	shell.contextMenu = menu;
	// `copy: false`: the capture phase reaches the host before the editing surface, so neither
	// the editor nor the browser copies.
	for (const type of ['copy', 'cut'] as const)
		element.addEventListener(
			type,
			(event) => {
				if (!core.viewOptions.customization.blockCopy) return;
				event.preventDefault();
				event.stopPropagation();
			},
			true,
		);

	const regionElement = (region: Region): HTMLElement | undefined =>
		region === 'ribbon'
			? shell.toolbar
			: region === 'status'
				? shell.chrome?.statusBar.element
				: undefined;
	const currentRegion = (): Region => {
		const active = activeElement();
		return (
			REGIONS.find((region) => active && regionElement(region)?.contains(active)) ?? 'document'
		);
	};
	const focusRegion = (region: Region) => {
		if (region === 'document') return toDocument();
		const root = regionElement(region);
		if (region === 'ribbon' && root) return (root as RibbonElement).focusTab();
		// The shared status bar draws its buttons in its shadow root; focus the first one on screen.
		const buttons = [
			...((root?.shadowRoot ?? root)?.querySelectorAll<HTMLElement>('button:not([hidden])') ?? []),
		];
		(buttons.find((button) => button.getClientRects().length > 0) ?? root)?.focus();
	};

	const layers: Layer[] = [
		{ element: undefined, isOpen: () => menu.isOpen, close: () => menu.close(), selfHandled: true },
		{ element: help.element, isOpen: () => help.isOpen, close: help.close, selfHandled: true },
		{
			element: core.inserts.linkDialog.element,
			isOpen: () => core.inserts.linkDialog.isOpen,
			close: () => core.inserts.linkDialog.close(),
			selfHandled: true,
		},
		{
			element: core.inserts.pictureDialog.element,
			isOpen: () => core.inserts.pictureDialog.isOpen,
			close: () => core.inserts.pictureDialog.close(),
			selfHandled: true,
		},
		...core.formatDialogs.list.map((dialog) => ({
			element: dialog.element,
			isOpen: () => dialog.isOpen,
			close: () => dialog.close(),
			selfHandled: true,
		})),
		{
			element: shell.chrome?.backstage.element,
			isOpen: () => Boolean(shell.chrome?.backstage.isOpen),
			close: () => shell.chrome?.closeBackstage(),
			selfHandled: true,
		},
		{
			element: shell.searchPanel?.element,
			isOpen: () => Boolean(shell.searchPanel?.isOpen),
			close: () => shell.searchPanel?.close(),
			selfHandled: true,
		},
		{
			element: shell.review?.commentsPanel.element,
			isOpen: () => Boolean(shell.review?.commentsOpen),
			close: () => {
				shell.review?.handleComments('toggle');
				toDocument();
			},
			selfHandled: false,
		},
	];
	const openLayer = () =>
		layers.find(
			(layer) =>
				layer.isOpen() &&
				(!presentation() || !layer.element || layer.element.getClientRects().length > 0),
		);
	// Snapshot before a layer's own handler runs. The presentation also needs Escape before
	// ProseMirror or native fullscreen consumes it.
	let escapeOwned = false;
	element.addEventListener(
		'keydown',
		(event) => {
			if (event.key !== 'Escape' || event.isComposing) return;
			const path = event.composedPath();
			escapeOwned = layers.some(
				(layer) =>
					layer.selfHandled &&
					layer.isOpen() &&
					(layer.element ? path.includes(layer.element) : path.includes(menu.element as Node)),
			);
			if (presentation() && !escapeOwned && !openLayer() && core.pages.views.escape()) {
				event.preventDefault();
				event.stopPropagation();
			}
		},
		true,
	);

	const actions: ShortcutActions = {
		find: () => {
			if (!presentation()) shell.searchPanel?.open('find');
		},
		replace: () => {
			if (!presentation()) shell.searchPanel?.open('replace');
		},
		save: () => void shell.chrome?.run('save'),
		print: () => void shell.chrome?.run('print'),
		moveRegion: (direction) => {
			const regions = presentation() ? (['document'] as const) : REGIONS;
			const current = currentRegion();
			const next =
				(regions.findIndex((region) => region === current) + direction + regions.length) %
				regions.length;
			const region = regions[next];
			if (region) focusRegion(region);
		},
		focusRibbon: () => {
			if (presentation()) return;
			focusRegion('ribbon');
			shell.toolbar?.dispatchEvent(new CustomEvent('dve-keytips'));
		},
		contextMenu: () => {
			if (currentRegion() === 'document') menu.openAtCaret();
		},
		help: () => {
			menu.close(false);
			help.open(
				shortcutRows(registry, mac(), core.viewOptions.customization.shortcuts),
				core.locale,
			);
		},
		escape: () => {
			if (escapeOwned) return false;
			const open = openLayer();
			if (open) open.close();
			else if (currentRegion() === 'document') return false;
			toDocument();
			return true;
		},
	};
	const mac = () => isMacPlatform();
	const registry = createShortcutRegistry(actions);
	installKeyboard(element, registry, () => ({
		readOnly: !core.canEditBody(),
		mac: mac(),
		disabled: core.viewOptions.customization.shortcuts,
	}));
}
