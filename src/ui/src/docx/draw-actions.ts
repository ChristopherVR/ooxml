// The Draw tab's commands and state for one editor. The pointer work (`draw-controller.ts`) loads on
// first use. Picking a drawing tool switches to Print Layout, as Word does, since ink floats on the
// pages.
import { DrawState } from '../ink/draw-state';
import { loadTools } from '../ink/tools';
import type { DocxDrawController } from './draw-controller';
import { defineDrawElements, syncDrawRibbon, type DrawAvailability } from './draw-ribbon';
import type { EditorCore } from './editor-core';
import { emit } from './events';
import { translate, type LocalizationKey } from './localization';
import type { RibbonAction } from './ribbon-action';

type DrawAction = Extract<RibbonAction, { type: 'draw' }>;

interface DrawSession {
	state: DrawState;
	controller?: DocxDrawController;
	loading?: Promise<DocxDrawController | undefined> | undefined;
}

const sessions = new WeakMap<EditorCore, DrawSession>();

function session(core: EditorCore): DrawSession {
	let current = sessions.get(core);
	if (!current) {
		current = { state: new DrawState() };
		const created = current;
		created.state.onChange(() => {
			created.controller?.sync();
			core.refreshControls();
		});
		sessions.set(core, current);
	}
	return current;
}

function report(core: EditorCore, error: Error): void {
	emit(core.element, 'document-error', error);
}

function controller(core: EditorCore): Promise<DocxDrawController | undefined> {
	const current = session(core);
	if (current.controller) return Promise.resolve(current.controller);
	current.loading ??= Promise.all([import('./draw-controller'), loadTools()]).then(
		([{ DocxDrawController }]) => {
			current.controller = new DocxDrawController({
				view: () => core.view,
				state: current.state,
				scroller: () => core.shell.canvas,
				pages: () => core.shell.printLayout?.element,
				active: () => core.canEditBody() && core.pages.viewMode === 'print',
				locale: () => core.locale,
				reportError: (error) => report(core, error),
				refresh: () => core.refreshControls(),
			});
			current.controller.sync();
			return current.controller;
		},
		(cause: unknown) => {
			current.loading = undefined;
			report(core, cause instanceof Error ? cause : new Error(String(cause)));
			return undefined;
		},
	);
	return current.loading;
}

/** Ink drawing happens on the Print Layout pages: switch there, as Word does. */
function toPages(core: EditorCore): void {
	if (core.pages.viewMode !== 'print') core.pages.setViewMode('print');
}

export function handleDrawAction(core: EditorCore, action: DrawAction): void {
	defineDrawElements();
	const { state } = session(core);
	switch (action.key) {
		case 'tool':
			if (action.value !== 'select') toPages(core);
			state.setTool(action.value);
			break;
		case 'pen':
			toPages(core);
			state.selectPen(action.index);
			break;
		case 'removePen':
			state.removePen(action.index);
			break;
		case 'addPen':
			toPages(core);
			if (!state.addPen(action.value))
				report(
					core,
					new Error(translate(core.locale, 'The pen gallery is full.' as LocalizationKey)),
				);
			break;
		case 'penColor':
			state.updatePen({ color: action.value });
			break;
		case 'thickness':
			state.updatePen({ width: action.value });
			break;
		case 'touch':
			state.setTouch(!state.snapshot.touch);
			break;
		case 'inkToShape':
			void controller(core).then((draw) => {
				draw?.inkToShape();
				keepRibbonFocus(core);
			});
			return;
		case 'inkToMath':
			return;
		case 'replay':
			toPages(core);
			void controller(core).then((draw) => draw?.replay());
			return;
	}
	if (state.canvasTool) void controller(core);
}

/**
 * Ink to Shape turns itself off once the selection is converted, which drops the focus it had:
 * the focus stays on the ribbon (on Lasso Select), where Ctrl+Z still undoes the conversion.
 */
function keepRibbonFocus(core: EditorCore): void {
	requestAnimationFrame(() => {
		const toolbar = core.shell.toolbar;
		const root = toolbar?.getRootNode() as ShadowRoot | Document | undefined;
		const active = root?.activeElement;
		const kept =
			active instanceof HTMLElement &&
			toolbar?.contains(active) &&
			!(active as HTMLButtonElement).disabled;
		if (toolbar && !kept)
			toolbar.querySelector<HTMLButtonElement>('button[data-draw="lasso"]')?.focus();
	});
}

/** Why each Draw command is unavailable now. */
function availability(core: EditorCore, draw: DocxDrawController | undefined): DrawAvailability {
	const editing = core.canEditBody() && Boolean(core.view?.editable);
	const reason = editing ? undefined : ('Editing is unavailable.' as const);
	return {
		draw: reason,
		inkToShape:
			reason ??
			(draw?.hasSelection() ? undefined : 'Select ink with Lasso Select or select an ink drawing.'),
		replaying: draw?.replaying ?? false,
		inkToMath: 'Ink to Math needs a handwriting recognizer, which this editor does not include.',
		replay:
			core.pages.viewMode === 'print' && !(draw ? draw.hasInk() || draw.replaying : hasInk(core))
				? 'There is no ink on the pages to replay.'
				: undefined,
	};
}

function hasInk(core: EditorCore): boolean {
	return Boolean(core.shell.printLayout?.element.querySelector('.dve-print-page [data-ink]'));
}

/** Refreshes the Draw tab and the drawing surface (after every transaction and layout). */
export function syncDraw(core: EditorCore): void {
	const current = sessions.get(core);
	const toolbar = core.shell.toolbar;
	const state = current?.state ?? new DrawState();
	current?.controller?.sync();
	if (toolbar) syncDrawRibbon(toolbar, state, availability(core, current?.controller), core.locale);
}

/** A new document: forget the strokes being joined, the lasso and a running replay. */
export function resetDraw(core: EditorCore): void {
	sessions.get(core)?.controller?.reset();
}

export function destroyDraw(core: EditorCore): void {
	sessions.get(core)?.controller?.destroy();
	sessions.delete(core);
}
