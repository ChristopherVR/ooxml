import type { ViewerState } from './controller.js';
import type { VsdxSource } from './contract.js';
import type { BackstagePage } from './backstage.js';

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
	announce(message: string): void;
	noteCount(): number;
}

/**
 * Visio's File backstage: opened from the File tab, closed with Back or Escape (focus returns to
 * File). Open, Save, Save As, Export, Print and Close are real; the rest is shown disabled.
 */
export class ViewerBackstage {
	readonly #root: HTMLElement;
	readonly #file: HTMLButtonElement;
	readonly #input: HTMLInputElement;
	#state: ViewerState | undefined;
	#urls = new Set<string>();
	constructor(private readonly host: BackstageHost) {
		this.#root = host.root.querySelector('.backstage')!;
		this.#file = host.root.querySelector('.file-tab')!;
		this.#input = this.#root.querySelector('.backstage-file-input')!;
	}
	get open(): boolean {
		return !this.#root.hidden;
	}
	show(page: BackstagePage = 'info'): void {
		this.#root.hidden = false;
		this.#file.setAttribute('aria-expanded', 'true');
		this.#select(page);
		this.#root.querySelector<HTMLElement>(`[data-backstage-item="${page}"]`)?.focus();
	}
	hide(): void {
		if (!this.open) return;
		this.#root.hidden = true;
		this.#file.setAttribute('aria-expanded', 'false');
		this.#file.focus();
	}
	wire(): () => void {
		const Abort = this.host.root.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.#file.addEventListener('click', () => (this.open ? this.hide() : this.show()), options);
		this.#root.addEventListener(
			'click',
			(event) => {
				const target = event.target as Element;
				if (target.closest?.('[data-backstage="back"]')) return this.hide();
				const item = target.closest?.<HTMLElement>('[data-backstage-item]')?.dataset.backstageItem;
				if (item === 'save') return this.#download();
				if (item === 'close') return this.#close();
				if (item) return this.#select(item as BackstagePage);
				const action = target.closest?.<HTMLButtonElement>('[data-backstage-action]');
				if (action && !action.disabled) this.#run(action.dataset.backstageAction!);
			},
			options,
		);
		this.#root.addEventListener(
			'keydown',
			(event) => {
				if (event.key !== 'Escape') return;
				event.preventDefault();
				event.stopPropagation();
				this.hide();
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
	#select(page: BackstagePage): void {
		for (const item of this.#root.querySelectorAll<HTMLElement>('[data-backstage-item]'))
			item.toggleAttribute('aria-current', item.dataset.backstageItem === page);
		for (const section of this.#root.querySelectorAll<HTMLElement>('[data-backstage-page]'))
			section.hidden = section.dataset.backstagePage !== page;
		if (page === 'print') this.#preview();
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
	/** Preview is a DOM clone of the rendered page, never document markup parsed from text. */
	#preview(): void {
		const target = this.#root.querySelector<HTMLElement>('.backstage-print-preview')!;
		const paper = this.host.viewport.querySelector('svg.paper');
		target.replaceChildren(...(paper ? [paper.cloneNode(true)] : []));
		const clone = target.querySelector('svg');
		clone?.removeAttribute('style');
		clone?.querySelectorAll('[tabindex]').forEach((node) => node.removeAttribute('tabindex'));
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
		for (const node of this.#root.querySelectorAll<HTMLButtonElement>(
			'[data-backstage-action="download"], [data-backstage-item="save"]',
		)) {
			node.disabled = !state.edit.sourceAvailable || busy;
			node.title = node.disabled ? saveReason : '';
		}
		for (const node of this.#root.querySelectorAll<HTMLButtonElement>(
			'[data-backstage-action="export-svg"], [data-backstage-action="print"]',
		))
			node.disabled = !page || busy;
		this.#root.querySelector<HTMLButtonElement>('[data-backstage-action="notes"]')!.disabled =
			notes === 0;
		this.#root.querySelector<HTMLButtonElement>('[data-backstage-item="close"]')!.disabled =
			!state.document;
	}
}
