import { findLocalizedControl } from './localization';
import { focusView } from './focus-view';
import type { DocumentView, OutlineAction, ViewPresentation } from './document-views';
import type { PageControllerHost } from './page-controller';
import { syncSelectTitles } from './ribbon';

type LayoutMode = Exclude<DocumentView, 'focus'>;

const loadView = {
	web: () => import('./view-web'),
	read: () => import('./view-read'),
	outline: () => import('./view-outline'),
	focus: () => import('./view-focus'),
};

/** Switches projections around the same live editor and keeps each layout's scroll position. */
export class ViewModes {
	mode: LayoutMode = 'draft';
	private presentation: ViewPresentation | undefined;
	private focus: ViewPresentation | undefined;
	private pending: DocumentView = 'draft';
	private sequence = 0;
	private readFrom: Exclude<LayoutMode, 'read'> = 'draft';
	private outlineFrom: Exclude<LayoutMode, 'outline'> = 'draft';
	private readonly scroll = new Map<LayoutMode, { top: number; left: number }>();

	constructor(private readonly host: PageControllerHost) {}

	get active(): DocumentView {
		return this.focus ? 'focus' : this.mode;
	}

	select(mode: DocumentView): void {
		this.request(mode, false);
	}

	private request(mode: DocumentView, restoring: boolean): void {
		if (mode === this.pending) return;
		this.pending = mode;
		const sequence = ++this.sequence;
		if (mode === this.mode || mode === 'draft' || mode === 'print') {
			this.apply(mode, undefined, restoring);
			return;
		}
		void loadView[mode]()
			.then((module) => {
				if (sequence !== this.sequence || !this.host.element.isConnected) return;
				this.apply(mode, module.mountView, restoring);
			})
			.catch((error: unknown) => {
				if (sequence !== this.sequence) return;
				this.pending = this.active;
				this.syncControls();
				this.host.reportError(error);
			});
	}

	escape(): boolean {
		// A view's own layer (Read Mode's note callout) closes before the view does.
		if (this.presentation?.escape?.()) return true;
		if (this.active === 'focus') this.select(this.mode);
		else if (this.mode === 'read') this.request(this.readFrom, true);
		else return false;
		return true;
	}

	run(action: OutlineAction): void {
		if (this.mode !== 'outline') return;
		if (action.key === 'close') this.request(this.outlineFrom, true);
		else this.presentation?.run?.(action);
	}

	sync(): void {
		const view = this.host.view();
		this.presentation?.sync?.(view);
		this.focus?.sync?.(view);
	}

	disconnect(): void {
		++this.sequence;
		this.pending = this.active;
		if (this.focus || this.mode === 'read' || this.mode === 'outline') {
			this.pending = 'draft';
			this.apply('draft');
		}
	}

	private apply(
		mode: DocumentView,
		mount?: (typeof import('./view-web'))['mountView'],
		restoring = false,
	): void {
		const { host } = this;
		const paper = host.paper();
		const canvas = paper?.parentElement;
		const frame = canvas?.closest<HTMLElement>('.dve-frame');
		if (!paper || !canvas || !frame) {
			this.pending = this.active;
			return;
		}
		const toolbar = host.toolbar();
		const previous = this.mode;
		const focusWasActive = Boolean(this.focus);
		const scrollBefore = { top: canvas.scrollTop, left: canvas.scrollLeft };
		this.focus?.dispose();
		this.focus = undefined;
		frame.removeAttribute('data-focus-view');
		if (mode === 'focus' && mount) {
			frame.setAttribute('data-focus-view', '');
			this.focus = mount({
				...host,
				frame,
				canvas,
				paper,
				toolbar,
				leave: () => this.select(this.mode),
			});
		} else if (mode !== 'focus' && mode !== previous) {
			host.beforeViewChange?.();
			this.scroll.set(previous, scrollBefore);
			this.presentation?.dispose();
			this.presentation = undefined;
			if (!restoring) {
				if (mode === 'read' && previous !== 'read') this.readFrom = previous;
				if (mode === 'outline' && previous !== 'outline') this.outlineFrom = previous;
			}
			this.mode = mode;
			canvas.dataset.documentView = mode;
			canvas.toggleAttribute('data-print-view', mode === 'print');
			host.printEditing?.(mode === 'print');
			const printLayout = host.printLayout();
			printLayout?.setActive(mode === 'print');
			paper.hidden = false;
			this.presentation = mount?.({
				...host,
				frame,
				canvas,
				paper,
				toolbar,
				leave: () => this.escape(),
			});
			host.configureRibbon?.();
			host.refreshEditable?.();
			if (mode === 'print' && printLayout) printLayout.relayoutNow(host.model());
		}
		this.syncControls();
		host.refreshControls();
		if (host.element.isConnected) {
			const view = host.view();
			if (this.mode === 'read' || !view?.editable) canvas.focus({ preventScroll: true });
			else focusView(view);
		}
		const scroll = mode === 'focus' || focusWasActive ? scrollBefore : this.scroll.get(this.mode);
		if (scroll && this.mode !== 'read') {
			canvas.scrollTop = scroll.top;
			canvas.scrollLeft = scroll.left;
		}
	}

	private syncControls(): void {
		const toolbar = this.host.toolbar();
		const select = toolbar && findLocalizedControl<HTMLSelectElement>(toolbar, 'Layout view');
		if (select) {
			select.value = this.active;
			syncSelectTitles(select.parentElement ?? toolbar);
		}
		this.host.statusBar()?.setViewMode(this.active);
	}
}
