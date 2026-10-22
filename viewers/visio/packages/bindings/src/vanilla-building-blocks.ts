import {
	buildingBlock,
	createBuildingBlockSession,
	effectiveMode,
	prepareDiagramBlock,
	type BuildingBlockOptions,
	type DiagramToolbarProps,
} from './building-blocks';
import {
	toolbarGroups,
	type ToolbarButton,
	type ToolbarControl,
	type ToolbarGroupId,
} from './toolbar-controls';

/** Replace children only when they differ, so a focused control is not moved by a refresh. */
function sync(parent: Element, children: readonly Node[]): void {
	const now = parent.childNodes;
	if (now.length === children.length && children.every((child, index) => now[index] === child))
		return;
	parent.replaceChildren(...children);
}
function attribute(element: Element, name: string, value: string | undefined): void {
	if (value === undefined) element.removeAttribute(name);
	else element.setAttribute(name, value);
}
/** Keyed DOM for the toolbar: elements are created once per control id and updated in place. */
function createToolbarView(doc: Document) {
	const elements = new Map<string, HTMLElement>();
	const element = <K extends keyof HTMLElementTagNameMap>(key: string, tag: K) => {
		let found = elements.get(key) as HTMLElementTagNameMap[K] | undefined;
		if (!found) {
			found = doc.createElement(tag);
			if (found instanceof HTMLButtonElement) found.type = 'button';
			elements.set(key, found);
		}
		return found;
	};
	const button = (control: ToolbarButton) => {
		const node = element(control.id, 'button');
		node.dataset.control = control.id;
		node.textContent = control.label;
		node.disabled = control.disabled;
		attribute(
			node,
			'aria-pressed',
			control.pressed === undefined ? undefined : String(control.pressed),
		);
		node.onclick = () => control.run();
		return node;
	};
	const render = (control: ToolbarControl): HTMLElement => {
		if (control.kind === 'button') return button(control);
		if (control.kind === 'text') {
			const node = element(control.id, 'span');
			node.dataset.control = control.id;
			node.textContent = control.text;
			return node;
		}
		const wrap = element('search', 'span');
		wrap.dataset.control = 'search';
		const input = element('search-input', 'input');
		input.type = 'search';
		input.setAttribute('aria-label', control.label);
		input.placeholder = control.label;
		if (input.value !== control.query) input.value = control.query;
		input.oninput = () => control.change(input.value);
		const summary = element('search-summary', 'span');
		summary.setAttribute('aria-live', 'polite');
		summary.textContent = control.summary;
		sync(wrap, [input, summary, button(control.previous), button(control.next)]);
		return wrap;
	};
	return (host: HTMLElement, props: DiagramToolbarProps, include?: readonly ToolbarGroupId[]) => {
		const groups = toolbarGroups(props, include).map((group) => {
			const span = element(`group-${group.id}`, 'span');
			span.dataset.group = group.id;
			sync(span, group.controls.map(render));
			return span;
		});
		sync(host, groups);
	};
}

/**
 * The vanilla building blocks: the same session the React, Vue, Solid, Svelte and Angular
 * adapters use, with `mount...` helpers in place of components. Mount the canvas, then any of the
 * optional pieces; each returns its disposer, and `destroy()` disposes whatever is still mounted.
 */
export function createViewerBuildingBlocks(initial: BuildingBlockOptions = {}) {
	// Permission is applied before anything mounts, so the canvas is never briefly editable.
	const session = createBuildingBlockSession(initial);
	let options = initial;
	session.update(options);
	const listeners = new Set<() => void>();
	const disposers = new Set<() => void>();
	const notify = () => {
		for (const listener of [...listeners]) listener();
	};
	const stopState = session.controller.subscribe(notify);
	const subscribe = (listener: () => void) => {
		listeners.add(listener);
		return () => {
			listeners.delete(listener);
		};
	};
	const track = (dispose: () => void) => {
		const once = () => {
			if (disposers.delete(once)) dispose();
		};
		disposers.add(once);
		return once;
	};
	const props = () => session.toolbar(options);
	return {
		controller: session.controller,
		/** Current flat control props: page, zoom, selection, search, enabled flags and callbacks. */
		get toolbarProps(): DiagramToolbarProps {
			return props();
		},
		get state() {
			return session.controller.state;
		},
		get loading() {
			return session.controller.state.loading;
		},
		get error() {
			return session.controller.state.error ?? session.controller.state.edit.error;
		},
		/** Effective permission: host `canEdit` and a drawing that can be saved back. */
		get mode() {
			return effectiveMode(session.controller.state);
		},
		/** Apply new options (content, `canEdit`, locale, callbacks). */
		update(next: BuildingBlockOptions) {
			options = next;
			session.update(next);
			notify();
		},
		/** Call `listener` on every state or option change; returns the unsubscribe function. */
		subscribe,
		/** Acknowledge a host save: clears dirty and fires `onDirtyChange(false)`. */
		markClean: () => session.markClean(),
		/** Mount the drawing surface; it follows `canEdit`. */
		mountCanvas(host: HTMLElement) {
			session.mount(host);
			return track(() => session.mount(null));
		},
		/** Mount the optional library toolbar; `include` picks groups. */
		mountToolbar(host: HTMLElement, include?: readonly ToolbarGroupId[]) {
			prepareDiagramBlock(host, 'toolbar');
			host.setAttribute('role', 'toolbar');
			const view = createToolbarView(host.ownerDocument);
			const draw = () => view(host, props(), include);
			const stop = subscribe(draw);
			draw();
			return track(() => {
				stop();
				host.replaceChildren();
			});
		},
		/** Mount the optional page selector. */
		mountPageNavigation(host: HTMLSelectElement) {
			prepareDiagramBlock(host, 'tabs');
			const draw = () => {
				const { pages, pageIndex, labels } = props();
				host.setAttribute('aria-label', labels.Page);
				sync(
					host,
					pages.map((page, index) => {
						const found = host.options[index];
						const option = found && found.dataset.pageId === page.id ? found : new Option();
						option.dataset.pageId = page.id;
						option.value = String(index);
						option.textContent = page.name;
						return option;
					}),
				);
				host.value = String(pageIndex);
			};
			host.onchange = () => props().onPageChange(Number(host.value));
			const stop = subscribe(draw);
			draw();
			return track(() => {
				stop();
				host.onchange = null;
				host.replaceChildren();
			});
		},
		/** Mount the optional read-only selection inspector. */
		mountInspector(host: HTMLElement) {
			return track(buildingBlock('mountDiagramInspector')(host, session.controller));
		},
		/** Mount the optional Shapes (stencil) pane. */
		mountShapes(host: HTMLElement) {
			return track(
				buildingBlock('mountDiagramShapes')(host, session.controller, { locale: options.locale }),
			);
		},
		destroy() {
			for (const dispose of [...disposers]) dispose();
			listeners.clear();
			stopState();
			session.destroy();
		},
	};
}
export type VanillaBuildingBlocks = ReturnType<typeof createViewerBuildingBlocks>;
export type {
	BuildingBlockOptions,
	BuildingBlockHostCallbacks,
	DiagramToolbarProps,
} from './building-blocks';
export type { ToolbarGroupId } from './toolbar-controls';
