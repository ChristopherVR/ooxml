import type { ViewerController, ViewerState } from './controller';
import {
	editErrorMessage,
	isEditCancellation,
	visioPageInsertCommand,
	visioQuarterTurnCommand,
	visioLocalRotationShape,
} from 'ooxml-core/visio/ui';
import { RIBBON_ACTION_EVENT, type VisioRibbonAction, type CanvasTool } from './ribbon-action';
import type { RibbonCommand } from './ribbon-parts';
import { routeRibbonAction, type RibbonTargets } from './ribbon-router';
import { ShapeDrawTool } from './viewer-draw-tool';
import type { Rulers } from './viewer-ruler';
import { ViewerPageOrder } from './viewer-page-order';
import { ViewerPageRename } from './viewer-page-rename';
import { ViewerPageDelete } from './viewer-page-delete';

export type { CanvasTool } from './ribbon-action';
interface CommandHost {
	root: ShadowRoot;
	viewport: HTMLElement;
	rulers: Rulers;
	controller: ViewerController;
	fit(mode: 'page' | 'width'): void;
	togglePane(pane: 'shapes' | 'inspector'): void;
	reveal(panel: 'edit' | 'notes' | 'selection' | 'layers', focusText: boolean): void;
	focusSearch(): void;
	togglePanZoom(): void;
	/** Transient command feedback for the status bar; document text is never interpreted as markup. */
	announce(message: string): void;
	toolChanged?(): void;
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
	#draw: ShapeDrawTool;
	#pageOrder: ViewerPageOrder;
	#pageRename: ViewerPageRename;
	#pageDelete: ViewerPageDelete;
	readonly #targets: RibbonTargets;
	constructor(private readonly host: CommandHost) {
		this.#pageOrder = new ViewerPageOrder(host.root, host.controller);
		this.#pageRename = new ViewerPageRename(host.root, host.controller);
		this.#pageDelete = new ViewerPageDelete(host.root, host.controller);
		this.#draw = new ShapeDrawTool(host.viewport, host.controller, {
			tool: () => (this.#tool === 'pointer' ? undefined : this.#tool),
			announce: host.announce,
		});
		this.#targets = {
			controller: host.controller,
			history: (key) => this.#history(key),
			deleteSelection: () => this.#delete(),
			rotateSelection: (direction) => this.#transform({ type: 'rotate', direction }),
			flipSelection: (axis) => this.#transform({ type: 'flip', axis }),
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
			togglePanZoom: host.togglePanZoom,
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
				if ((event as CustomEvent<{ command?: unknown }>).detail?.command === 'reorder-pages')
					this.#pageOrder.show();
				if ((event as CustomEvent<{ command?: unknown }>).detail?.command === 'rename-page')
					this.#pageRename.show();
				if ((event as CustomEvent<{ command?: unknown }>).detail?.command === 'delete-page')
					this.#pageDelete.show();
				if (
					(event as CustomEvent<{ command?: unknown }>).detail?.command === 'tab-add' &&
					(event.target as Element)?.matches?.('.page-tabs')
				)
					this.#insertPage();
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
			this.#pageOrder.close();
			this.#pageRename.close();
			this.#pageDelete.close();
			++this.#pending;
			events.abort();
			disposeDraw();
		};
	}
	setTool(tool: CanvasTool): void {
		if (tool !== 'pointer' && !this.#canEdit(this.host.controller.state)) return;
		this.#tool = tool;
		this.render(this.host.controller.state);
		this.host.toolChanged?.();
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
	#insertPage(): void {
		const state = this.host.controller.state;
		const page = state.document?.pages[state.pageIndex];
		if (!page || !state.document || !this.#canEdit(state)) return;
		void this.#edit(async () => {
			const command = visioPageInsertCommand(state.document!, page.id);
			await this.host.controller.applyEdits([command]);
			const index = this.host.controller.state.document!.pages.findIndex(
				(page) => page.id === command.pageId,
			);
			if (index >= 0) this.host.controller.setPage(index);
		}, 'Inserted a blank page.');
	}
	#transform(action: Extract<VisioRibbonAction, { type: 'rotate' | 'flip' }>): void {
		const state = this.host.controller.state;
		const page = state.document?.pages[state.pageIndex];
		if (!page || !state.selectedShape || !this.#canEdit(state)) return;
		const command =
			action.type === 'rotate'
				? visioQuarterTurnCommand(page, state.selectedShape.id, action.direction)
				: visioLocalRotationShape(page, state.selectedShape.id)
					? {
							type: 'flip-shape' as const,
							pageId: page.id,
							shapeId: state.selectedShape.id,
							axis: action.axis,
						}
					: undefined;
		if (command)
			void this.#edit(
				() => this.host.controller.applyEdits([command]),
				action.type === 'rotate' ? `Rotated ${action.direction} 90°.` : `Flipped ${action.axis}.`,
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
		if (control && key === '9') return { type: 'tool', tool: 'ellipse' };
		if (control && key === '6') return { type: 'tool', tool: 'line' };
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
		this.#pageOrder.render(state);
		this.#pageRename.render(state);
		this.#pageDelete.render(state);
		const { root, viewport } = this.host;
		const button = (name: string) => root.querySelector<RibbonCommand>(`[command="${name}"]`)!;
		const box = (name: string) =>
			root.querySelector<RibbonCommand & { checked: boolean }>(`[data-check="${name}"]`)!;
		const editing = this.#canEdit(state);
		if (this.#tool !== 'pointer' && !state.edit.sourceAvailable) this.#tool = 'pointer';
		const page = state.document?.pages[state.pageIndex];
		const rotating =
			editing &&
			!!page &&
			!!state.selectedShape &&
			!!visioLocalRotationShape(page, state.selectedShape.id);
		for (const name of ['rotate-left', 'rotate-right', 'flip-horizontal', 'flip-vertical'])
			button(name).disabled = !rotating;
		root.querySelector<RibbonCommand>('[data-menu="rotate"]')!.disabled = !rotating;
		root.querySelector<RibbonCommand>('[data-menu="position"]')!.disabled = !rotating;
		button('undo').disabled = !state.edit.canUndo || state.edit.busy || state.loading;
		button('redo').disabled = !state.edit.canRedo || state.edit.busy || state.loading;
		button('pointer').setAttribute('pressed', String(this.#tool === 'pointer'));
		// The drawing-tools split button shows the active tool; its active drawing item is checked.
		const rectangle = button('rectangle');
		rectangle.toggleAttribute('data-active', this.#tool === 'rectangle');
		rectangle.disabled = !editing || !page;
		button('rectangle-item').setAttribute('checked', String(this.#tool === 'rectangle'));
		button('rectangle-item').disabled = !editing || !page;
		button('line-tool').setAttribute('checked', String(this.#tool === 'line'));
		button('ellipse').setAttribute('checked', String(this.#tool === 'ellipse'));
		button('ellipse').disabled = !editing || !page;
		button('line-tool').disabled = !editing || !page;
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
