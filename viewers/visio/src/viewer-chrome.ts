import type { VisioDocument } from 'ooxml-core/visio';
import type { ViewerController, ViewerState } from './controller.js';
import { editControlsTemplate } from './viewer-edit-controls.js';
import { layerControlsTemplate } from './viewer-layer-controls.js';
import { searchTemplate } from './viewer-search.js';

/** Only static application markup belongs here. All document labels are inserted as text. */
export const viewerChromeTemplate = `
<div class="toolbar" role="group" aria-label="Diagram controls">

  <div class="ribbon-tabs" role="tablist" aria-label="Ribbon">
    <button type="button" id="home-tab" role="tab" aria-selected="true" aria-controls="home-panel" data-tab="home">Home</button>
    <button type="button" id="view-tab" role="tab" aria-selected="false" aria-controls="view-panel" tabindex="-1" data-tab="view">View</button>
  </div>
  <div class="ribbon-content" id="home-panel" role="tabpanel" aria-labelledby="home-tab">
    <details class="ribbon-tools" open><summary>Tools <span aria-hidden="true">⌄</span></summary><div class="tools-content"><div class="ribbon-group search-group">${searchTemplate}<span class="group-label">Find in drawing</span></div>
    <div class="ribbon-group"><div class="ribbon-actions">
      <button class="ribbon-command" type="button" data-chrome="selection"><span class="command-icon" data-icon="selection" aria-hidden="true"></span><span>Shape details</span></button>
      <button class="ribbon-command" type="button" data-chrome="edit"><span class="command-icon" data-icon="edit" aria-hidden="true"></span><span>Edit text</span></button>
      <button class="ribbon-command" type="button" data-chrome="layers"><span class="command-icon" data-icon="layers" aria-hidden="true"></span><span>Layers</span></button>
    </div><span class="group-label">Inspect and edit</span></div>
    </div></details>
  <div class="ribbon-primary">
    <label class="page-picker">Page <select aria-label="Page"></select></label>
    <div class="page-stepper"><button type="button" data-chrome="previous-page" aria-label="Previous page">‹</button><button type="button" data-chrome="next-page" aria-label="Next page">›</button></div>
    <span class="spacer"></span><span class="local-label">Local workspace</span>
  </div>
    <p class="ribbon-hint">Select a shape to inspect its details.<br>Text editing is experimental.</p>
  </div>
  <div class="ribbon-content" id="view-panel" role="tabpanel" aria-labelledby="view-tab" hidden>
    <div class="ribbon-group"><div class="ribbon-actions">
      <button class="ribbon-command" type="button" data-chrome="pages" aria-pressed="true" aria-controls="page-rail"><span class="command-icon" data-icon="pages" aria-hidden="true"></span><span>Pages pane</span></button>
      <button class="ribbon-command" type="button" data-chrome="inspector" aria-pressed="true" aria-controls="inspector-pane"><span class="command-icon" data-icon="inspector" aria-hidden="true"></span><span>Inspector pane</span></button>
      <button class="ribbon-command" type="button" data-chrome="notes"><span class="command-icon" data-icon="notes" aria-hidden="true"></span><span>Review notes</span></button>
    </div><span class="group-label">Workspace panes</span></div>
    <p class="ribbon-hint">Canvas shortcuts: + / − to zoom, 0 to fit.<br>Arrow keys move between focused shapes.</p>
  </div>
</div>
<div class="workspace">
  <nav id="page-rail" class="page-rail" aria-label="Diagram pages"><div class="pane-heading"><span>Pages</span><span data-page-count>0</span></div><ol class="page-list"></ol><p class="page-empty">Open a drawing to see its pages.</p></nav>
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
</div>
<div class="status"><div class="status-message" role="status"><span data-status></span><span data-diagnostics></span></div><div class="notes-strip"><button type="button" data-chrome="notes" aria-controls="inspector-pane" aria-expanded="false">Notes<span class="notes-count" data-note-count></span></button><span>Files stay in your browser</span></div><slot name="workspace-footer"></slot><div class="zoom-controls" role="group" aria-label="Canvas zoom"><button type="button" data-action="fit">Fit page</button><button type="button" data-action="actual">100%</button><span class="zoom-divider"></span><button type="button" data-action="out" aria-label="Zoom out">−</button><output class="zoom" aria-label="Zoom level">100%</output><button type="button" data-action="in" aria-label="Zoom in">+</button></div></div>`;

/** Presentation state stays local; page navigation uses the same controller as every binding. */
export class ViewerChrome {
	#root: ShadowRoot;
	#controller: ViewerController;
	#document: VisioDocument | null | undefined;
	#pageList: HTMLOListElement;
	#rail: HTMLElement;
	#inspector: HTMLElement;
	#notes: HTMLDetailsElement;
	#responsive: MediaQueryList | undefined;
	#compact: MediaQueryList | undefined;
	#pagesManuallyToggled = false;
	#inspectorManuallyToggled = false;
	constructor(root: ShadowRoot, controller: ViewerController) {
		this.#root = root;
		this.#controller = controller;
		this.#pageList = root.querySelector('.page-list')!;
		this.#rail = root.querySelector('.page-rail')!;
		this.#inspector = root.querySelector('.inspector-pane')!;
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
		const tools = this.#root.querySelector<HTMLDetailsElement>('.ribbon-tools')!;
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
			if (!this.#pagesManuallyToggled) this.#setPane('pages', !this.#responsive?.matches);
			if (!this.#inspectorManuallyToggled) this.#setPane('inspector', !compact?.matches);
		};
		responsive();
		this.#responsive?.addEventListener('change', responsive, options);
		compact?.addEventListener('change', responsive, options);
		this.#root.addEventListener(
			'click',
			(event) => {
				const button = (event.target as Element)?.closest?.<HTMLButtonElement>('button');
				if (!button || button.disabled) return;
				if (compact?.matches && button.dataset.chrome) tools.open = false;
				const tab = button.dataset.tab;
				if (tab === 'home' || tab === 'view') this.#showTab(tab);
				if (button.dataset.pageIndex !== undefined)
					this.#controller.setPage(Number(button.dataset.pageIndex));
				const action = button.dataset.chrome;
				if (action === 'previous-page')
					this.#controller.setPage(this.#controller.state.pageIndex - 1);
				if (action === 'next-page') this.#controller.setPage(this.#controller.state.pageIndex + 1);
				if (action === 'pages') {
					this.#pagesManuallyToggled = true;
					this.#setPane('pages', this.#rail.hidden);
				}
				if (action === 'inspector') {
					this.#inspectorManuallyToggled = true;
					this.#setPane('inspector', this.#inspector.hidden);
				}
				if (
					action === 'notes' ||
					action === 'selection' ||
					action === 'edit' ||
					action === 'layers'
				)
					this.#reveal(action);
			},
			options,
		);
		this.#root.addEventListener(
			'keydown',
			(event) => {
				const key = event as KeyboardEvent;
				const button = (event.target as Element)?.closest?.<HTMLButtonElement>('button');
				if (
					!button ||
					!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(key.key)
				)
					return;
				if (button.dataset.tab) {
					if (key.key === 'ArrowUp' || key.key === 'ArrowDown') return;
					key.preventDefault();
					const tab =
						key.key === 'Home'
							? 'home'
							: key.key === 'End'
								? 'view'
								: button.dataset.tab === 'home'
									? 'view'
									: 'home';
					this.#showTab(tab);
					this.#root.querySelector<HTMLButtonElement>(`[data-tab="${tab}"]`)!.focus();
				} else if (button.dataset.pageIndex !== undefined) {
					key.preventDefault();
					const pages = [...this.#pageList.querySelectorAll<HTMLButtonElement>('button')];
					const current = Number(button.dataset.pageIndex);
					const next =
						key.key === 'Home'
							? 0
							: key.key === 'End'
								? pages.length - 1
								: Math.max(
										0,
										Math.min(
											pages.length - 1,
											current + (['ArrowRight', 'ArrowDown'].includes(key.key) ? 1 : -1),
										),
									);
					this.#controller.setPage(next);
					pages[next]?.focus();
				}
			},
			options,
		);
		this.#notes.addEventListener('toggle', () => this.#syncNotes(), options);
		return () => events.abort();
	}
	#showTab(tab: 'home' | 'view'): void {
		for (const button of this.#root.querySelectorAll<HTMLButtonElement>('[data-tab]')) {
			const selected = button.dataset.tab === tab;
			button.setAttribute('aria-selected', String(selected));
			button.tabIndex = selected ? 0 : -1;
		}
		for (const panel of this.#root.querySelectorAll<HTMLElement>('[role="tabpanel"]'))
			panel.hidden = panel.id !== `${tab}-panel`;
	}
	#setPane(pane: 'pages' | 'inspector', visible: boolean): void {
		if (visible && this.#compact?.matches) {
			const other = pane === 'pages' ? 'inspector' : 'pages';
			(other === 'pages' ? this.#rail : this.#inspector).hidden = true;
			this.#root.querySelector(`[data-chrome="${other}"]`)!.setAttribute('aria-pressed', 'false');
		}
		(pane === 'pages' ? this.#rail : this.#inspector).hidden = !visible;
		this.#root
			.querySelector(`[data-chrome="${pane}"]`)!
			.setAttribute('aria-pressed', String(visible));
		this.#syncNotes();
	}
	#syncNotes(): void {
		this.#root
			.querySelector('.notes-strip button')!
			.setAttribute('aria-expanded', String(!this.#inspector.hidden && this.#notes.open));
	}
	#reveal(kind: 'notes' | 'selection' | 'edit' | 'layers'): void {
		this.#inspectorManuallyToggled = true;
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
		panel.querySelector('summary')!.focus();
		panel.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
		this.#syncNotes();
	}
	render(state: ViewerState, noteCount: number): void {
		const page = state.document?.pages[state.pageIndex];
		if (state.document !== this.#document) {
			this.#document = state.document;
			const focused = this.#root.activeElement as HTMLElement | null;
			const restoreFocus = focused?.dataset.pageIndex;
			const fragment = this.#root.ownerDocument.createDocumentFragment();
			for (const [index, candidate] of (state.document?.pages ?? []).entries()) {
				const item = this.#root.ownerDocument.createElement('li');
				const button = this.#root.ownerDocument.createElement('button');
				button.type = 'button';
				button.className = 'page-link';
				button.dataset.pageIndex = String(index);
				button.setAttribute('aria-label', `Go to page ${index + 1}: ${candidate.name}`);
				const number = this.#root.ownerDocument.createElement('span');
				number.className = 'page-number';
				number.textContent = String(index + 1);
				const card = this.#root.ownerDocument.createElement('span');
				card.className = 'page-card';
				const mark = this.#root.ownerDocument.createElement('span');
				mark.className = 'page-mark';
				mark.textContent = '▤';
				mark.setAttribute('aria-hidden', 'true');
				const name = this.#root.ownerDocument.createElement('span');
				name.className = 'page-name';
				name.textContent = candidate.name;
				const meta = this.#root.ownerDocument.createElement('span');
				meta.className = 'page-meta';
				meta.textContent = candidate.isBackground
					? 'Background page'
					: `${candidate.width} × ${candidate.height} in`;
				card.append(mark, name, meta);
				button.append(number, card);
				item.append(button);
				fragment.append(item);
			}
			this.#pageList.replaceChildren(fragment);
			if (restoreFocus !== undefined)
				[...this.#pageList.querySelectorAll<HTMLButtonElement>('button')]
					.find((button) => button.dataset.pageIndex === restoreFocus)
					?.focus();
		}
		for (const button of this.#pageList.querySelectorAll<HTMLButtonElement>('button')) {
			const selected = Number(button.dataset.pageIndex) === state.pageIndex;
			if (selected) button.setAttribute('aria-current', 'page');
			else button.removeAttribute('aria-current');
			button.tabIndex = selected ? 0 : -1;
		}
		const count = state.document?.pages.length ?? 0;
		this.#root.querySelector('[data-page-count]')!.textContent = String(count);
		this.#root.querySelector<HTMLElement>('.page-empty')!.hidden = count > 0;
		this.#root.querySelector('[data-page-name]')!.textContent = page?.name ?? 'No diagram open';
		this.#root.querySelector('[data-page-size]')!.textContent = page
			? `${page.width} × ${page.height} in`
			: 'No page';
		this.#root.querySelector('[data-page-shapes]')!.textContent = page
			? `${page.shapes.length} top-level`
			: 'No shapes';
		this.#root.querySelector<HTMLElement>('.selection-hint')!.hidden = !!state.selectedShape;
		this.#root.querySelector<HTMLButtonElement>('[data-chrome="previous-page"]')!.disabled =
			!page || state.pageIndex === 0;
		this.#root.querySelector<HTMLButtonElement>('[data-chrome="next-page"]')!.disabled =
			!page || state.pageIndex >= count - 1;
		this.#root.querySelector<HTMLButtonElement>('[data-chrome="selection"]')!.disabled =
			!state.selectedShape;
		this.#root.querySelector<HTMLButtonElement>('[data-chrome="edit"]')!.disabled = !page;
		this.#root.querySelector<HTMLButtonElement>('[data-chrome="layers"]')!.disabled =
			this.#root.querySelector<HTMLDetailsElement>('.layer-controls')!.hidden;
		for (const button of this.#root.querySelectorAll<HTMLButtonElement>('[data-chrome="notes"]'))
			button.disabled = noteCount === 0;
		this.#root.querySelector('[data-note-count]')!.textContent = noteCount ? ` · ${noteCount}` : '';
		this.#syncNotes();
	}
}
