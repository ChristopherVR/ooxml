import type { ViewerController, ViewerState } from './controller.js';
import { editErrorMessage, isEditCancellation } from './edit-error.js';
import { RIBBON_ACTION_EVENT, type VisioRibbonAction } from './ribbon-action.js';
import type { RibbonCommand } from './ribbon-parts.js';
import { routeRibbonAction, type RibbonTargets } from './ribbon-router.js';
import { RectangleDrawTool } from './viewer-draw-tool.js';
import type { Rulers } from './viewer-ruler.js';

export type CanvasTool = 'pointer' | 'rectangle';
interface CommandHost {
	root: ShadowRoot;
	viewport: HTMLElement;
	rulers: Rulers;
	controller: ViewerController;
	fit(mode: 'page' | 'width'): void;
	togglePane(pane: 'shapes' | 'inspector'): void;
	reveal(panel: 'edit' | 'notes' | 'selection' | 'layers', focusText: boolean): void;
	focusSearch(): void;
	/** Transient command feedback for the status bar; document text is never interpreted as markup. */
	announce(message: string): void;
}
const editable = (target: EventTarget | null) =>
	target instanceof Element &&
	!!target.closest('input, textarea, select, [contenteditable]:not([contenteditable="false"])');

/**
 * Visio command state (tool, grid, pending edits) and keyboard shortcuts. Ribbon controls emit
 * typed `ribbon-action` events and shortcuts build the same actions; both go through
 * `routeRibbonAction`. Document mutation stays in the controller and core.
 */
export class ViewerCommands {
	#tool: CanvasTool = 'pointer';
	#grid = false;
	#ruler = false;
	#pending = 0;
	#draw: RectangleDrawTool;
	readonly #targets: RibbonTargets;
	constructor(private readonly host: CommandHost) {
		this.#draw = new RectangleDrawTool(host.viewport, host.controller, {
			active: () => this.#tool === 'rectangle',
			announce: host.announce,
		});
		this.#targets = {
			controller: host.controller,
			history: (key) => this.#history(key),
			deleteSelection: () => this.#delete(),
			setTool: (tool) => this.setTool(tool),
			toggleGrid: () => {
				this.#grid = !this.#grid;
				this.render(host.controller.state);
			},
			toggleRuler: () => {
				this.#ruler = !this.#ruler;
				this.render(host.controller.state);
			},
			toggleFullscreen: () => this.#toggleFullscreen(),
			togglePane: host.togglePane,
			reveal: host.reveal,
			fit: host.fit,
			focusSearch: host.focusSearch,
		};
	}
	get tool(): CanvasTool {
		return this.#tool;
	}
	/** Run one typed action, exactly as a ribbon control or shortcut would. */
	run(action: VisioRibbonAction): void {
		routeRibbonAction(this.#targets, action);
	}
	wire(): () => void {
		const { root, viewport } = this.host;
		const doc = root.ownerDocument;
		const Abort = doc.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		root.addEventListener(
			RIBBON_ACTION_EVENT,
			(event) => this.run((event as CustomEvent<VisioRibbonAction>).detail),
			options,
		);
		// The shared zoom slider's fit button is the status bar's Fit page to current window.
		root.addEventListener(
			'office-command',
			(event) => {
				if ((event as CustomEvent<{ command?: unknown }>).detail?.command === 'zoom-fit')
					this.run({ type: 'zoom', mode: 'fit' });
			},
			options,
		);
		root.addEventListener('keydown', (event) => this.#shortcut(event as KeyboardEvent), options);
		viewport.addEventListener(
			'wheel',
			(event) => {
				if (!event.ctrlKey && !event.metaKey) return;
				event.preventDefault();
				const zoom = this.host.controller.state.zoom;
				this.host.controller.setZoom(event.deltaY < 0 ? zoom * 1.1 : zoom / 1.1);
			},
			{ ...options, passive: false },
		);
		// Like Visio's drawing window, a canvas click takes keyboard focus so shortcuts apply.
		viewport.addEventListener(
			'pointerdown',
			() => {
				if (!viewport.contains(root.activeElement)) viewport.focus({ preventScroll: true });
			},
			options,
		);
		viewport.addEventListener(
			'dblclick',
			(event) => {
				if (this.#tool !== 'pointer') return;
				if ((event.target as Element)?.closest?.('[data-shape-id]'))
					this.run({ type: 'reveal', panel: 'edit', focusText: true });
			},
			options,
		);
		doc.addEventListener(
			'fullscreenchange',
			() => this.render(this.host.controller.state),
			options,
		);
		const disposeDraw = this.#draw.wire();
		return () => {
			++this.#pending;
			events.abort();
			disposeDraw();
		};
	}
	setTool(tool: CanvasTool): void {
		if (tool === 'rectangle' && !this.#canEdit(this.host.controller.state)) return;
		this.#tool = tool;
		this.render(this.host.controller.state);
	}
	#canEdit(state: ViewerState): boolean {
		return state.edit.sourceAvailable && !state.loading && !state.edit.busy;
	}
	#history(key: 'undo' | 'redo'): void {
		const { controller } = this.host;
		const { edit, loading } = controller.state;
		const available = key === 'undo' ? edit.canUndo : edit.canRedo;
		if (available && !edit.busy && !loading) void this.#edit(() => controller[key]());
	}
	#delete(): void {
		const state = this.host.controller.state;
		const shape = state.selectedShape;
		const pageId = shape?.pageId ?? state.document?.pages[state.pageIndex]?.id;
		if (!shape || pageId === undefined || !this.#canEdit(state)) return;
		void this.#edit(
			() => this.host.controller.applyEdits([{ type: 'delete-shape', pageId, shapeId: shape.id }]),
			`Deleted ${shape.name || `shape ${shape.id}`}.`,
		);
	}
	async #edit(action: () => Promise<void>, success?: string): Promise<void> {
		const request = ++this.#pending;
		const { root, viewport } = this.host;
		const canvasFocused = viewport.contains(root.activeElement);
		try {
			await action();
			// A re-rendered page drops the focused shape; keep shortcuts on the drawing window.
			if (canvasFocused && !viewport.contains(root.activeElement))
				viewport.focus({ preventScroll: true });
			if (success && request === this.#pending) this.host.announce(success);
		} catch (error) {
			// The controller records refused edits in state.edit.error for the status bar.
			if (
				request === this.#pending &&
				!isEditCancellation(error) &&
				!this.host.controller.state.edit.error
			)
				this.host.announce(editErrorMessage(error));
		}
	}
	#toggleFullscreen(): void {
		const host = this.host.root.host as HTMLElement;
		const doc = this.host.root.ownerDocument;
		const request =
			doc.fullscreenElement === host ? doc.exitFullscreen?.() : host.requestFullscreen?.();
		request?.catch?.(() =>
			this.host.announce('Full screen is not available in this browser frame.'),
		);
	}
	/** Visio shortcuts as typed actions. Text fields keep native undo, deletion and editing keys. */
	#shortcutAction(event: KeyboardEvent): VisioRibbonAction | undefined {
		const state = this.host.controller.state;
		const control = event.ctrlKey || event.metaKey;
		const key = event.key.length === 1 ? event.key.toLowerCase() : event.key;
		if (control && key === 'f') return { type: 'search' };
		if (control && event.shiftKey && key === 'w') return { type: 'zoom', mode: 'fit' };
		if (control && (key === 'PageDown' || key === 'PageUp'))
			return { type: 'page', step: key === 'PageDown' ? 1 : -1 };
		if (key === 'F5' && !control) return { type: 'fullscreen' };
		if (editable(event.target)) return undefined;
		if (control && !event.shiftKey && key === 'z') return { type: 'history', key: 'undo' };
		if (control && (key === 'y' || (event.shiftKey && key === 'z')))
			return { type: 'history', key: 'redo' };
		if (control && key === '1') return { type: 'tool', tool: 'pointer' };
		if (control && key === '8') return { type: 'tool', tool: 'rectangle' };
		if (!control && key === 'Delete' && state.selectedShape) return { type: 'delete' };
		if (!control && key === 'F2' && state.document)
			return { type: 'reveal', panel: 'edit', focusText: true };
		if (key === 'Escape' && this.#tool !== 'pointer' && !this.#draw.drawing)
			return { type: 'tool', tool: 'pointer' };
		return undefined;
	}
	#shortcut(event: KeyboardEvent): void {
		if (event.defaultPrevented || event.isComposing || event.altKey) return;
		const action = this.#shortcutAction(event);
		if (!action) return;
		event.preventDefault();
		this.run(action);
	}
	render(state: ViewerState): void {
		const { root, viewport } = this.host;
		const button = (name: string) => root.querySelector<RibbonCommand>(`[command="${name}"]`)!;
		const box = (name: string) =>
			root.querySelector<RibbonCommand & { checked: boolean }>(`[data-check="${name}"]`)!;
		const editing = this.#canEdit(state);
		if (this.#tool === 'rectangle' && !state.edit.sourceAvailable) this.#tool = 'pointer';
		const page = state.document?.pages[state.pageIndex];
		button('undo').disabled = !state.edit.canUndo || state.edit.busy || state.loading;
		button('redo').disabled = !state.edit.canRedo || state.edit.busy || state.loading;
		button('pointer').setAttribute('pressed', String(this.#tool === 'pointer'));
		// The drawing-tools split button shows the active tool; its Rectangle item is checked.
		const rectangle = button('rectangle');
		rectangle.toggleAttribute('data-active', this.#tool === 'rectangle');
		rectangle.disabled = !editing || !page;
		button('rectangle-item').setAttribute('checked', String(this.#tool === 'rectangle'));
		button('rectangle-item').disabled = !editing || !page;
		rectangle.title = state.edit.sourceAvailable
			? 'Rectangle (Ctrl+8)'
			: 'Rectangle (Ctrl+8): open a .vsdx file to draw. Model-only documents are read only.';
		box('grid').checked = this.#grid;
		box('grid').disabled = !page;
		box('ruler').checked = this.#ruler;
		box('ruler').disabled = !page;
		for (const name of ['zoom-fit', 'page-width']) button(name).disabled = !page;
		root.querySelector<RibbonCommand>('[data-menu="zoom"]')!.disabled = !page;
		const fullscreen = button('fullscreen');
		const host = root.host as HTMLElement;
		fullscreen.disabled = typeof host.requestFullscreen !== 'function';
		fullscreen.setAttribute('pressed', String(root.ownerDocument.fullscreenElement === host));
		viewport.dataset.tool = this.#tool;
		viewport.dataset.grid = String(this.#grid);
		this.host.rulers.render(this.#ruler && !!page, state.zoom);
	}
}
