import type { ViewerState } from './controller.js';
import type { VsdxSource } from './contract.js';
import { backstageItems, type BackstagePage } from './backstage.js';

/** What the backstage needs from the element; every action delegates to existing APIs. */
export interface BackstageHost {
	root: ShadowRoot;
	viewport: HTMLElement;
	fileName(): string;
	load(source: VsdxSource): Promise<void>;
	exportVsdx(): { bytes: Uint8Array; dirty: boolean };
	exportSvg(): { svg: string; pageIndex: number };
	closeDocument(): void;
	revealNotes(): void;
	showOptions(): void;
	announce(message: string): void;
	noteCount(): number;
}

/**
 * Visio's File backstage: opened from the File tab, closed with Back or Escape (focus returns to
 * File). Open, Save, Save As, Export, Print and Close are real; the rest is shown disabled.
 */
type Backstage = HTMLElement & {
	items: unknown;
	open: boolean;
	show(page?: string): void;
	close(): void;
};
type Ribbon = HTMLElement & { focusFile(): void };

export class ViewerBackstage {
	readonly #root: Backstage;
	readonly #ribbon: Ribbon;
	readonly #input: HTMLInputElement;
	#state: ViewerState | undefined;
	#items = '';
	#urls = new Set<string>();
	constructor(private readonly host: BackstageHost) {
		this.#root = host.root.querySelector<Backstage>('office-ui-backstage')!;
		this.#ribbon = host.root.querySelector<Ribbon>('office-ui-ribbon')!;
		this.#input = this.#root.querySelector('.backstage-file-input')!;
	}
	get open(): boolean {
		return this.#root.open;
	}
	show(page: BackstagePage = 'info'): void {
		this.#ribbon.setAttribute('file-expanded', 'true');
		this.#root.show(page);
		if (page === 'print') this.#preview();
	}
	hide(): void {
		this.#root.close();
	}
	wire(): () => void {
		const Abort = this.host.root.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.#ribbon.addEventListener(
			'office-ribbon-file',
			() => (this.open ? this.hide() : this.show()),
			options,
		);
		this.#root.addEventListener(
			'office-backstage-select',
			(event) => {
				const id = (event as CustomEvent<{ id: string }>).detail.id;
				if (id === 'save') this.#download();
				else if (id === 'close') this.#close();
				else if (id === 'options') {
					// Visio opens its Options dialog over the drawing, not a backstage page.
					this.hide();
					this.host.showOptions();
				} else if (id === 'print') this.#preview();
			},
			options,
		);
		// Back, Escape or a command closed the backstage: return focus to File.
		this.#root.addEventListener(
			'office-backstage-close',
			() => {
				this.#ribbon.setAttribute('file-expanded', 'false');
				this.#ribbon.focusFile();
			},
			options,
		);
		this.#root.addEventListener(
			'click',
			(event) => {
				const action = (event.target as Element).closest?.<HTMLButtonElement>(
					'[data-backstage-action]',
				);
				if (action && !action.disabled) this.#run(action.dataset.backstageAction!);
			},
			options,
		);
		this.#input.addEventListener(
			'change',
			() => {
				const file = this.#input.files?.[0];
				this.#input.value = '';
				if (!file) return;
				this.hide();
				this.host
					.load(file)
					.catch((error: unknown) =>
						this.host.announce(error instanceof Error ? error.message : String(error)),
					);
			},
			options,
		);
		return () => {
			events.abort();
			for (const url of this.#urls) URL.revokeObjectURL(url);
			this.#urls.clear();
		};
	}
	#run(action: string): void {
		if (action === 'open') this.#input.click();
		else if (action === 'download') this.#download();
		else if (action === 'export-svg') this.#exportSvg();
		else if (action === 'print') this.#print();
		else if (action === 'notes') {
			this.hide();
			this.host.revealNotes();
		}
	}
	#save(blob: Blob, name: string): void {
		const doc = this.host.root.ownerDocument;
		const url = URL.createObjectURL(blob);
		this.#urls.add(url);
		const anchor = doc.createElement('a');
		anchor.href = url;
		anchor.download = name;
		anchor.hidden = true;
		this.host.root.append(anchor);
		anchor.click();
		anchor.remove();
		// Keep the URL alive through browser download dispatch, then release it.
		doc.defaultView?.setTimeout(() => {
			if (this.#urls.delete(url)) URL.revokeObjectURL(url);
		}, 1000);
	}
	#base(): string {
		return this.host.fileName().replace(/\.vsdx?$/i, '') || 'Drawing';
	}
	#download(): void {
		if (!this.#state?.edit.sourceAvailable) return;
		try {
			const result = this.host.exportVsdx();
			this.#save(
				new Blob([new Uint8Array(result.bytes)], { type: 'application/vnd.ms-visio.drawing' }),
				`${this.#base()}-${result.dirty ? 'edited' : 'original'}-copy.vsdx`,
			);
			this.hide();
			this.host.announce(
				'VSDX copy download requested. Native Visio compatibility is not verified.',
			);
		} catch (error) {
			this.host.announce(error instanceof Error ? error.message : String(error));
		}
	}
	#exportSvg(): void {
		try {
			const result = this.host.exportSvg();
			this.#save(
				new Blob([result.svg], { type: 'image/svg+xml;charset=utf-8' }),
				`${this.#base()}-page-${result.pageIndex + 1}.svg`,
			);
			this.hide();
			this.host.announce('SVG export requested. This is an approximate snapshot.');
		} catch (error) {
			this.host.announce(error instanceof Error ? error.message : String(error));
		}
	}
	/** The shared print preview shows an inert clone of the rendered page, never parsed markup. */
	#preview(): void {
		const preview = this.#root.querySelector<HTMLElement & { pages: Node[] }>(
			'office-ui-print-preview',
		)!;
		const paper = this.host.viewport.querySelector('svg.paper');
		preview.pages = paper ? [paper] : [];
	}
	#print(): void {
		const doc = this.host.root.ownerDocument;
		const paper = this.host.viewport.querySelector('svg.paper');
		if (!paper || !doc.body) return;
		const frame = doc.createElement('iframe');
		frame.setAttribute('aria-hidden', 'true');
		frame.style.cssText = 'position:fixed;width:0;height:0;border:0;visibility:hidden';
		doc.body.append(frame);
		const target = frame.contentDocument;
		const view = frame.contentWindow;
		if (!target || !view) return frame.remove();
		const clone = target.importNode(paper, true) as SVGSVGElement;
		clone.removeAttribute('style');
		clone.setAttribute('width', '100%');
		target.body.style.margin = '0';
		target.body.append(clone);
		view.addEventListener('afterprint', () => frame.remove(), { once: true });
		view.focus();
		view.print();
	}
	#close(): void {
		if (!this.#state?.document) return;
		this.hide();
		this.host.closeDocument();
	}
	render(state: ViewerState): void {
		this.#state = state;
		const page = state.document?.pages[state.pageIndex];
		const set = (key: string, value: string) => {
			const node = this.#root.querySelector(`[data-info="${key}"]`);
			if (node) node.textContent = value;
		};
		set('name', state.document ? this.host.fileName() || 'Untitled drawing' : 'No drawing open');
		set('pages', String(state.document?.pages.length ?? 0));
		set('page', page?.name ?? 'None');
		set('size', page ? `${page.width} × ${page.height} in` : 'None');
		set('shapes', page ? `${page.shapes.length} top-level` : 'None');
		set(
			'state',
			state.document?.format === 'vsd'
				? 'Legacy VSD preview (read only)'
				: !state.edit.sourceAvailable
					? state.document
						? 'Model-only preview (read only)'
						: 'None'
					: state.edit.dirty
						? 'Edited copy (not saved)'
						: 'Original',
		);
		const notes = this.host.noteCount();
		this.#root.querySelector('.backstage-note-count')!.textContent = notes
			? `${notes} compatibility notes describe what this viewer approximates or omits.`
			: 'No compatibility notes.';
		const busy = state.loading || state.edit.busy;
		const saveReason =
			state.document?.format === 'vsd'
				? 'Legacy VSD drawings are read only here.'
				: 'Open a .vsdx file to save a copy.';
		const saveDisabled = !state.edit.sourceAvailable || busy;
		const download = this.#root.querySelector<HTMLButtonElement>(
			'[data-backstage-action="download"]',
		)!;
		download.disabled = saveDisabled;
		download.title = saveDisabled ? saveReason : '';
		// Rebuild the navigation only when Save or Close change, so focus stays put.
		const items = backstageItems({
			save: { disabled: saveDisabled, ...(saveDisabled ? { title: saveReason } : {}) },
			close: { disabled: !state.document },
		});
		const key = JSON.stringify(items);
		if (key !== this.#items) {
			this.#items = key;
			this.#root.items = items;
		}
		for (const node of this.#root.querySelectorAll<HTMLButtonElement>(
			'[data-backstage-action="export-svg"], [data-backstage-action="print"]',
		))
			node.disabled = !page || busy;
		this.#root.querySelector<HTMLButtonElement>('[data-backstage-action="notes"]')!.disabled =
			notes === 0;
	}
}
