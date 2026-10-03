import type { VisioDocument } from 'ooxml-core/visio';
import type { OfficeTab } from 'ooxml-ui/controls';
import type { ViewerController, ViewerState } from './controller.js';
import { editControlsTemplate } from './viewer-edit-controls.js';
import { layerControlsTemplate } from './viewer-layer-controls.js';

/**
 * Static workspace markup (legacy; migrate to builders when next changed). The ribbon, page
 * tabs and status bar are built by `createRibbon`, `createPageTabs` and `createStatusBar`.
 * All document labels are inserted as text.
 */
export const viewerChromeTemplate = `<div class="workspace">
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
</div>`;

/** Presentation state stays local; page navigation uses the same controller as every binding. */
export class ViewerChrome {
	#root: ShadowRoot;
	#controller: ViewerController;
	#document: VisioDocument | null | undefined;
	#pageList: HTMLOListElement;
	#pageTabs: HTMLElement & { tabs: OfficeTab[]; selected: string };
	#tools: HTMLDetailsElement;
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
		this.#pageTabs = root.querySelector('office-ui-tab-strip')!;
		this.#tools = root.querySelector('.ribbon-tools')!;
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
				const tab = button.dataset.tab;
				if (tab) this.showTab(tab);
				if (button.dataset.pageIndex !== undefined)
					this.#controller.setPage(Number(button.dataset.pageIndex));
				const action = button.dataset.chrome;
				if (action === 'inspector') this.togglePane('inspector');
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
				} else if (button.dataset.pageIndex !== undefined) {
					key.preventDefault();
					const pages = [
						...this.#pageList.querySelectorAll<HTMLButtonElement>('[data-page-index]'),
					];
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
	showTab(tab: string): void {
		for (const button of this.#root.querySelectorAll<HTMLButtonElement>('[data-tab]')) {
			const selected = button.dataset.tab === tab;
			button.setAttribute('aria-selected', String(selected));
			button.tabIndex = selected ? 0 : -1;
		}
		for (const panel of this.#root.querySelectorAll<HTMLElement>('[role="tabpanel"]'))
			panel.hidden = panel.id !== `${tab}-panel`;
	}
	/** On phones, close the Tools sheet once a command runs. */
	closeCompactTools(): void {
		if (this.#compact?.matches) this.#tools.open = false;
	}
	/** Show or hide a task pane; compact layouts close the Tools menu afterwards. */
	togglePane(pane: 'pages' | 'inspector'): void {
		if (this.#compact?.matches) this.#tools.open = false;
		if (pane === 'pages') this.#pagesManuallyToggled = true;
		else this.#inspectorManuallyToggled = true;
		this.#setPane(pane, (pane === 'pages' ? this.#rail : this.#inspector).hidden);
	}
	/** A ribbon command or menu item by its stable id (the first match wins). */
	#command(name: string): HTMLElement & { disabled: boolean } {
		return this.#root.querySelector(`[command="${name}"]`)!;
	}
	#setPane(pane: 'pages' | 'inspector', visible: boolean): void {
		if (visible && this.#compact?.matches) {
			const other = pane === 'pages' ? 'inspector' : 'pages';
			(other === 'pages' ? this.#rail : this.#inspector).hidden = true;
			this.#command(other).setAttribute('checked', 'false');
		}
		(pane === 'pages' ? this.#rail : this.#inspector).hidden = !visible;
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
		const text = panel.querySelector<HTMLTextAreaElement>('textarea');
		if (focusText && text && !text.disabled) text.focus();
		else panel.querySelector('summary')!.focus();
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
			// The shared tab strip inserts labels as text.
			this.#pageTabs.tabs = (state.document?.pages ?? []).map((candidate, index) => ({
				id: String(index),
				label: candidate.name,
				title: candidate.isBackground ? `${candidate.name} (background page)` : candidate.name,
			}));
			if (restoreFocus !== undefined)
				[...this.#pageList.querySelectorAll<HTMLButtonElement>('button')]
					.find((button) => button.dataset.pageIndex === restoreFocus)
					?.focus();
		}
		this.#pageTabs.selected = page ? String(state.pageIndex) : '';
		for (const button of this.#pageList.querySelectorAll<HTMLButtonElement>('button')) {
			const selected = Number(button.dataset.pageIndex) === state.pageIndex;
			if (selected) button.setAttribute('aria-current', 'page');
			else button.removeAttribute('aria-current');
			button.tabIndex = selected ? 0 : -1;
		}
		this.#root
			.querySelector('[data-page-status]')!
			.setAttribute(
				'value',
				page ? `Page ${state.pageIndex + 1} of ${state.document!.pages.length}` : '',
			);
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
