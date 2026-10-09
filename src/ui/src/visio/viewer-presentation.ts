import type { VisioDocument, VisioPage } from 'ooxml-core/visio';
import type { ViewerController, ViewerState } from './controller';
import { renderPage } from './render-svg';
import { mapPresentationKey, type PresentationKeyBuffer } from './presentation-keys';

/** Visio presents foreground pages only; background pages appear under the pages using them. */
export function presentationPages(document: VisioDocument | null | undefined): VisioPage[] {
	return document?.pages.filter((page) => !page.isBackground) ?? [];
}

function barButton(doc: Document, name: string, label: string, icon: string): HTMLButtonElement {
	const button = doc.createElement('button');
	button.type = 'button';
	button.dataset.present = name;
	button.setAttribute('aria-label', label);
	button.title = label;
	const glyph = doc.createElement('office-ui-icon');
	glyph.setAttribute('name', icon);
	button.append(glyph);
	return button;
}

/** The chrome-free presentation surface: a stage, a page indicator and previous/next/exit. */
export function createPresentation(doc: Document): HTMLElement {
	const overlay = doc.createElement('div');
	overlay.className = 'presentation';
	overlay.hidden = true;
	overlay.tabIndex = -1;
	overlay.setAttribute('role', 'dialog');
	overlay.setAttribute('aria-modal', 'true');
	overlay.setAttribute('aria-label', 'Presentation Mode');
	const stage = doc.createElement('div');
	stage.className = 'presentation-stage';
	const bar = doc.createElement('div');
	bar.className = 'presentation-bar';
	bar.setAttribute('role', 'toolbar');
	bar.setAttribute('aria-label', 'Presentation controls');
	const indicator = doc.createElement('span');
	indicator.className = 'presentation-page';
	indicator.setAttribute('aria-live', 'polite');
	bar.append(
		barButton(doc, 'previous', 'Previous page', 'chevronLeft'),
		indicator,
		barButton(doc, 'next', 'Next page', 'chevronRight'),
		barButton(doc, 'exit', 'Exit Presentation Mode (Esc)', 'close'),
	);
	overlay.append(stage, bar);
	return overlay;
}

/**
 * Visio's Presentation Mode (F5): the current foreground page fitted to the screen on a dark
 * backdrop, read only, one page at a time. It renders its own copy of each page, so the drawing
 * window's page, zoom and selection are never touched and are exactly as before on exit. The
 * Fullscreen API is used when the browser allows it; otherwise the overlay covers the window.
 */
export class ViewerPresentation {
	readonly overlay: HTMLElement;
	#stage: HTMLElement;
	#indicator: HTMLElement;
	#active = false;
	#fullscreen = false;
	#index = 0;
	#document: VisioDocument | null = null;
	#layers: ViewerState['layerVisibilityOverrides'] | undefined;
	#shown: VisioPage | undefined;
	#dispose: () => void = () => {};
	#buffer: PresentationKeyBuffer = { digits: '' };
	#inerted: Element[] = [];
	#focus: HTMLElement | null = null;
	#notice = '';
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
	) {
		this.overlay = createPresentation(root.ownerDocument);
		root.append(this.overlay);
		this.#stage = this.overlay.querySelector('.presentation-stage')!;
		this.#indicator = this.overlay.querySelector('.presentation-page')!;
	}
	get active(): boolean {
		return this.#active;
	}
	/** 0-based position among the foreground pages, while presenting. */
	get index(): number {
		return this.#index;
	}
	start(): void {
		const state = this.controller.state;
		const pages = presentationPages(state.document);
		if (this.#active || !pages.length || state.loading) return;
		const current = state.document!.pages[state.pageIndex];
		this.#active = true;
		this.#buffer.digits = '';
		this.#index = Math.max(0, current ? pages.indexOf(current) : 0);
		const active = this.root.activeElement ?? this.root.ownerDocument.activeElement;
		this.#focus = active instanceof HTMLElement ? active : null;
		this.#inerted = [...this.root.children].filter(
			(child) =>
				child !== this.overlay &&
				!(child instanceof HTMLStyleElement) &&
				!child.hasAttribute('inert'),
		);
		for (const child of this.#inerted) child.setAttribute('inert', '');
		this.overlay.hidden = false;
		this.overlay.dataset.fallback = 'true';
		this.render(state);
		this.overlay.focus({ preventScroll: true });
		const request = this.overlay.requestFullscreen;
		if (typeof request !== 'function') return;
		try {
			void Promise.resolve(request.call(this.overlay)).then(
				() => {
					if (this.#active) delete this.overlay.dataset.fallback;
				},
				() => {},
			);
		} catch {
			// Frames without fullscreen permission keep the in-element overlay.
		}
	}
	exit(): void {
		if (!this.#active) return;
		this.#active = false;
		this.#buffer.digits = '';
		const doc = this.root.ownerDocument;
		if (this.#fullscreen || this.root.fullscreenElement === this.overlay) {
			this.#fullscreen = false;
			void Promise.resolve(doc.exitFullscreen?.()).catch(() => {});
		}
		this.overlay.hidden = true;
		delete this.overlay.dataset.fallback;
		this.#dispose();
		this.#dispose = () => {};
		this.#stage.replaceChildren();
		this.#shown = undefined;
		this.#document = null;
		for (const child of this.#inerted) child.removeAttribute('inert');
		this.#inerted = [];
		const focus = this.#focus;
		this.#focus = null;
		(focus?.isConnected ? focus : this.root.querySelector<HTMLElement>('.viewport'))?.focus({
			preventScroll: true,
		});
		this.render(this.controller.state);
	}
	/** Move to a 0-based foreground page, clamped. */
	go(index: number): void {
		if (!this.#active) return;
		this.#notice = '';
		const count = presentationPages(this.controller.state.document).length;
		this.#index = Math.max(0, Math.min(count - 1, index));
		this.render(this.controller.state);
	}
	wire(): () => void {
		const doc = this.root.ownerDocument;
		const Abort = doc.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.overlay.addEventListener('keydown', (event) => this.#key(event), options);
		this.overlay.addEventListener(
			'click',
			(event) => {
				const control = (event.target as Element | null)?.closest?.('[data-present]');
				const name = (control as HTMLElement | null)?.dataset.present;
				if (name === 'exit') return this.exit();
				if (name === 'previous') return this.go(this.#index - 1);
				if (name === 'next' || !(event.target as Element | null)?.closest?.('.presentation-bar'))
					this.go(this.#index + 1);
			},
			options,
		);
		this.overlay.addEventListener('contextmenu', (event) => event.preventDefault(), options);
		this.root.addEventListener(
			'click',
			(event) => {
				if ((event.target as Element | null)?.closest?.('.presentation-mode')) this.start();
			},
			options,
		);
		doc.addEventListener(
			'fullscreenchange',
			() => {
				if (this.root.fullscreenElement === this.overlay) {
					// A request that resolves after the show ended leaves full screen at once.
					if (!this.#active) return void Promise.resolve(doc.exitFullscreen?.()).catch(() => {});
					this.#fullscreen = true;
					delete this.overlay.dataset.fallback;
				} else if (this.#fullscreen) {
					// The browser left full screen (its own Esc, F11 or a tab switch): end the show.
					this.#fullscreen = false;
					this.exit();
				}
			},
			options,
		);
		return () => {
			events.abort();
			this.exit();
		};
	}
	#key(event: KeyboardEvent): void {
		// The drawing window's shortcuts (edits, tools, key tips) never run while presenting.
		event.stopPropagation();
		this.#notice = '';
		const onControl = (event.target as Element | null)?.closest?.('[data-present]');
		if (onControl && (event.key === 'Enter' || event.key === ' ')) return;
		const action = mapPresentationKey(event, this.#buffer);
		if (action.type === 'none') {
			// F5 would reload a host page; Visio keeps presenting.
			if (event.key === 'F5') event.preventDefault();
			this.#renderIndicator();
			return;
		}
		event.preventDefault();
		const count = presentationPages(this.controller.state.document).length;
		if (action.type === 'exit') return this.exit();
		if (action.type === 'next') return this.go(this.#index + 1);
		if (action.type === 'previous') return this.go(this.#index - 1);
		if (action.type === 'first') return this.go(0);
		if (action.type === 'last') return this.go(count - 1);
		if (action.type === 'goto') {
			if (action.page <= count) return this.go(action.page - 1);
			this.#notice = `No page ${action.page} (${count} pages)`;
		}
		this.#renderIndicator();
	}
	/** Sync the entry points and, while presenting, the shown page with the controller state. */
	render(state: ViewerState): void {
		const pages = presentationPages(state.document);
		const unavailable = !pages.length || state.loading;
		for (const button of this.root.querySelectorAll<HTMLElement & { disabled: boolean }>(
			'.presentation-mode, [command="presentation"]',
		)) {
			button.disabled = unavailable;
			button.title = unavailable
				? 'Presentation Mode: open a drawing with a foreground page.'
				: 'Presentation Mode (F5)';
		}
		if (!this.#active) return;
		if (!pages.length) return this.exit();
		if (state.document !== this.#document && this.#shown) {
			// After an edit or replacement, keep showing the same page when it still exists.
			const same = pages.findIndex((page) => page.id === this.#shown!.id);
			if (same >= 0) this.#index = same;
		}
		this.#index = Math.max(0, Math.min(pages.length - 1, this.#index));
		const page = pages[this.#index]!;
		if (
			page !== this.#shown ||
			state.document !== this.#document ||
			state.layerVisibilityOverrides !== this.#layers
		) {
			this.#dispose();
			this.#dispose = () => {};
			this.#document = state.document;
			this.#layers = state.layerVisibilityOverrides;
			this.#shown = page;
			const result = renderPage(state.document!, page, {
				interactive: false,
				layerVisibilityOverrides: state.layerVisibilityOverrides,
			});
			this.#dispose = result.dispose;
			const ratio = page.width / page.height;
			result.svg.style.aspectRatio = `${page.width} / ${page.height}`;
			result.svg.style.width = `min(100cqw, ${+(ratio * 100).toFixed(4)}cqh)`;
			this.#stage.replaceChildren(result.svg);
		}
		this.#renderIndicator();
	}
	#renderIndicator(): void {
		const pages = presentationPages(this.controller.state.document);
		const page = pages[this.#index];
		if (!this.#active || !page) return;
		this.#indicator.textContent = this.#notice
			? this.#notice
			: this.#buffer.digits
				? `Go to page ${this.#buffer.digits}`
				: `${page.name} (${this.#index + 1} of ${pages.length})`;
		this.#indicator.title = page.name;
		const button = (name: string) =>
			this.overlay.querySelector<HTMLButtonElement>(`[data-present="${name}"]`)!;
		button('previous').disabled = this.#index === 0;
		button('next').disabled = this.#index === pages.length - 1;
	}
}
