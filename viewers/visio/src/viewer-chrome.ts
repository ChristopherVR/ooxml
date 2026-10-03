import type { VisioDocument } from 'ooxml-core/visio';
import type { OfficeTab } from 'ooxml-ui/controls';
import type { ViewerController, ViewerState } from './controller.js';
import { editControlsTemplate } from './viewer-edit-controls.js';
import { layerControlsTemplate } from './viewer-layer-controls.js';

/**
 * Static workspace markup (legacy; migrate to builders when next changed). The ribbon, Shapes
 * window, page tabs and status bar are built by their own modules. Visio keeps pages in the
 * bottom tabs and the All pages list, so there is no page pane. Labels are inserted as text.
 */
export const viewerChromeTemplate = `<div class="workspace">
  <div class="viewport" tabindex="0" role="region" aria-label="Diagram canvas"></div>
  <aside id="inspector-pane" class="inspector-pane" aria-label="Drawing inspector">
    <div class="pane-heading"><span>Inspector</span><span class="inspector-kind">Drawing</span><button class="pane-close" type="button" data-chrome="inspector" aria-label="Close inspector">×</button></div>
    <div class="inspector-body">
      <section class="inspector-card document-card" aria-label="Current page"><h2>Current page</h2><p class="current-page-name" data-page-name>No diagram open</p><dl><dt>Size</dt><dd data-page-size>No page</dd><dt>Shapes</dt><dd data-page-shapes>No shapes</dd></dl></section>
      <p class="selection-hint">Select a shape on the canvas to see its data and links.</p>
      <details class="shape-inspector inspector-card" hidden><summary>Selected shape</summary><div></div></details>
      ${editControlsTemplate}${layerControlsTemplate}
      <details class="notes inspector-card"><summary>Compatibility notes</summary><ul></ul></details>
    </div>
  </aside>
</div>`;

export type TaskPane = 'shapes' | 'inspector';

/** Presentation state stays local; page navigation uses the same controller as every binding. */
export class ViewerChrome {
	#root: ShadowRoot;
	#controller: ViewerController;
	#document: VisioDocument | null | undefined;
	#pageTabs: HTMLElement & { tabs: OfficeTab[]; selected: string };
	#allPages: HTMLElement & { disabled: boolean };
	#tools: HTMLDetailsElement;
	#inspector: HTMLElement;
	#panes: Record<TaskPane, HTMLElement>;
	#notes: HTMLDetailsElement;
	#responsive: MediaQueryList | undefined;
	#compact: MediaQueryList | undefined;
	/** Panes the user showed or hid; responsive defaults leave them alone afterwards. */
	#manual = new Set<TaskPane>();
	constructor(root: ShadowRoot, controller: ViewerController) {
		this.#root = root;
		this.#controller = controller;
		this.#pageTabs = root.querySelector('office-ui-tab-strip')!;
		this.#allPages = root.querySelector('[data-menu="all-pages"]')!;
		this.#tools = root.querySelector('.ribbon-tools')!;
		this.#inspector = root.querySelector('.inspector-pane')!;
		this.#panes = { shapes: root.querySelector('.shapes-pane')!, inspector: this.#inspector };
		this.#notes = root.querySelector('.notes')!;
	}
	wire(): () => void {
		const view = this.#root.ownerDocument.defaultView;
		const Abort = view?.AbortController ?? AbortController;
		const events = new Abort(),
			options = { signal: events.signal };
		this.#responsive = view?.matchMedia?.('(max-width: 980px)');
		const compact = view?.matchMedia?.('(max-width: 760px)');
		this.#compact = compact;
		const tools = this.#tools;
		const toolLayout = () => {
			tools.open = !compact?.matches;
		};
		toolLayout();
		compact?.addEventListener('change', toolLayout, options);
		this.#root.addEventListener(
			'keydown',
			(event) => {
				if ((event as KeyboardEvent).key === 'Escape' && compact?.matches && tools.open) {
					event.preventDefault();
					event.stopPropagation();
					tools.open = false;
					tools.querySelector<HTMLElement>('summary')!.focus();
				}
			},
			{ ...options, capture: true },
		);
		const responsive = () => {
			// Visio's defaults: the Shapes window on wide screens, the inspector unless on a phone.
			if (!this.#manual.has('shapes')) this.#setPane('shapes', !this.#responsive?.matches);
			if (!this.#manual.has('inspector')) this.#setPane('inspector', !compact?.matches);
		};
		responsive();
		this.#responsive?.addEventListener('change', responsive, options);
		compact?.addEventListener('change', responsive, options);
		this.#root.addEventListener(
			'click',
			(event) => {
				const button = (event.target as Element)?.closest?.<HTMLButtonElement>('button');
				if (!button || button.disabled) return;
				const tab = button.dataset.tab;
				if (tab) this.showTab(tab);
				const action = button.dataset.chrome;
				if (action === 'inspector' || action === 'shapes') this.togglePane(action);
				if (action === 'notes') this.reveal('notes');
			},
			options,
		);
		this.#root.addEventListener(
			'office-tab-select',
			(event) => {
				const id = Number((event as CustomEvent<{ id: string }>).detail.id);
				if (Number.isSafeInteger(id)) this.#controller.setPage(id);
			},
			options,
		);
		// Visio's All pages list beside the page tabs.
		this.#allPages.addEventListener(
			'office-command',
			(event) => {
				const command = (event as CustomEvent<{ command: string }>).detail.command;
				const index = /^page-(\d+)$/.exec(command)?.[1];
				if (index === undefined) return;
				event.stopPropagation();
				this.#controller.setPage(Number(index));
			},
			options,
		);
		this.#root.addEventListener(
			'keydown',
			(event) => {
				const key = event as KeyboardEvent;
				const button = (event.target as Element)?.closest?.<HTMLButtonElement>('button');
				if (!button?.dataset.tab || !['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(key.key))
					return;
				key.preventDefault();
				const tabs = [...this.#root.querySelectorAll<HTMLButtonElement>('[data-tab]')];
				const index = tabs.indexOf(button);
				const next =
					key.key === 'Home'
						? tabs[0]
						: key.key === 'End'
							? tabs.at(-1)
							: tabs[(index + (key.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length];
				if (!next?.dataset.tab) return;
				this.showTab(next.dataset.tab);
				next.focus();
			},
			options,
		);
		this.#notes.addEventListener('toggle', () => this.#syncNotes(), options);
		return () => events.abort();
	}
	showTab(tab: string): void {
		for (const button of this.#root.querySelectorAll<HTMLButtonElement>('[data-tab]')) {
			const selected = button.dataset.tab === tab;
			button.setAttribute('aria-selected', String(selected));
			button.tabIndex = selected ? 0 : -1;
		}
		for (const panel of this.#root.querySelectorAll<HTMLElement>('.ribbon-content'))
			panel.hidden = panel.id !== `${tab}-panel`;
	}
	/** On phones, close the Tools sheet once a command runs. */
	closeCompactTools(): void {
		if (this.#compact?.matches) this.#tools.open = false;
	}
	/** Show or hide a task pane; compact layouts close the Tools menu afterwards. */
	togglePane(pane: TaskPane): void {
		if (this.#compact?.matches) this.#tools.open = false;
		this.#manual.add(pane);
		this.#setPane(pane, this.#panes[pane].hidden);
	}
	/** A ribbon command or menu item by its stable id (the first match wins). */
	#command(name: string): HTMLElement & { disabled: boolean } {
		return this.#root.querySelector(`[command="${name}"]`)!;
	}
	#setPane(pane: TaskPane, visible: boolean): void {
		// Phones show one pane at a time.
		if (visible && this.#compact?.matches)
			for (const other of Object.keys(this.#panes) as TaskPane[])
				if (other !== pane) {
					this.#panes[other].hidden = true;
					this.#command(other).setAttribute('checked', 'false');
				}
		this.#panes[pane].hidden = !visible;
		this.#command(pane).setAttribute('checked', String(visible));
		this.#syncNotes();
	}
	#syncNotes(): void {
		this.#root
			.querySelector('.notes-strip button')!
			.setAttribute('aria-expanded', String(!this.#inspector.hidden && this.#notes.open));
	}
	/** Open an inspector disclosure; text editing may focus the text field directly (F2). */
	reveal(kind: 'notes' | 'selection' | 'edit' | 'layers', focusText = false): void {
		if (this.#compact?.matches) this.#tools.open = false;
		this.#manual.add('inspector');
		this.#setPane('inspector', true);
		const selector =
			kind === 'notes'
				? '.notes'
				: kind === 'edit'
					? '.edit-controls'
					: kind === 'layers'
						? '.layer-controls'
						: '.shape-inspector';
		const panel = this.#root.querySelector<HTMLDetailsElement>(selector)!;
		if (panel.hidden) return;
		panel.open = true;
		const text = panel.querySelector<HTMLTextAreaElement>('textarea');
		if (focusText && text && !text.disabled) text.focus();
		else panel.querySelector('summary')!.focus();
		panel.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
		this.#syncNotes();
	}
	render(state: ViewerState, noteCount: number): void {
		const page = state.document?.pages[state.pageIndex];
		const doc = this.#root.ownerDocument;
		if (state.document !== this.#document) {
			this.#document = state.document;
			const pages = state.document?.pages ?? [];
			// The shared tab strip and menu items insert labels as text.
			this.#pageTabs.tabs = pages.map((candidate, index) => ({
				id: String(index),
				label: candidate.name,
				title: candidate.isBackground ? `${candidate.name} (background page)` : candidate.name,
			}));
			this.#allPages.replaceChildren(
				...pages.map((candidate, index) => {
					const item = doc.createElement('office-ui-menu-item');
					item.setAttribute('command', `page-${index}`);
					item.setAttribute(
						'label',
						candidate.isBackground ? `${candidate.name} (background)` : candidate.name,
					);
					return item;
				}),
			);
		}
		this.#pageTabs.selected = page ? String(state.pageIndex) : '';
		for (const item of this.#allPages.querySelectorAll('office-ui-menu-item'))
			item.setAttribute(
				'checked',
				String(item.getAttribute('command') === `page-${state.pageIndex}`),
			);
		this.#allPages.disabled = !page;
		this.#root
			.querySelector('[data-page-status]')!
			.setAttribute(
				'value',
				page ? `Page ${state.pageIndex + 1} of ${state.document!.pages.length}` : '',
			);
		this.#root.querySelector('[data-page-name]')!.textContent = page?.name ?? 'No diagram open';
		this.#root.querySelector('[data-page-size]')!.textContent = page
			? `${page.width} × ${page.height} in`
			: 'No page';
		this.#root.querySelector('[data-page-shapes]')!.textContent = page
			? `${page.shapes.length} top-level`
			: 'No shapes';
		this.#root.querySelector<HTMLElement>('.selection-hint')!.hidden = !!state.selectedShape;
		const layers = this.#root.querySelector<HTMLDetailsElement>('.layer-controls')!.hidden;
		this.#command('shape-data').disabled = !state.selectedShape;
		this.#command('layer-properties').disabled = layers;
		this.#command('layers-pane').disabled = layers;
		this.#command('notes').disabled = noteCount === 0;
		this.#root.querySelector<HTMLButtonElement>('.notes-strip button')!.disabled = noteCount === 0;
		this.#root.querySelector('[data-note-count]')!.textContent = noteCount ? ` · ${noteCount}` : '';
		this.#syncNotes();
	}
}
