import { command, combo, type CommandSpec, type RibbonCommand } from './ribbon-parts';
import type { VisioRibbonAction } from './ribbon-action';

type Combo = HTMLElement & {
	options: { value: string; label: string; disabled?: boolean }[];
	value: string;
	disabled: boolean;
};

/** Where the toolbar sits relative to the point it is shown for, and the gap to keep from edges. */
const GAP = 8;
const MARGIN = 8;

interface Mirror {
	/** The toolbar's own command, and the Home ribbon command whose state it shows. */
	mini: string;
	ribbon: string;
}

/**
 * Visio's Mini Toolbar: the text formatting commands of Home (font, size, grow, shrink, bold,
 * italic, underline, alignment) in a floating bar beside the shape you right-click or the text you
 * select. Its buttons emit the same typed actions as the ribbon's, so formatting is applied by the
 * one code path; their state (pressed, disabled) and the font and size lists are copied from the
 * ribbon's own controls after every render, never worked out a second time. It never takes focus,
 * so a text edit in progress stays where it is.
 */
export class ViewerMiniToolbar {
	readonly element: HTMLElement;
	readonly #ribbon: HTMLElement;
	readonly #mirrors: Mirror[] = [];
	#enabled = true;
	constructor(root: ShadowRoot, ribbon: HTMLElement) {
		this.#ribbon = ribbon;
		const doc = root.ownerDocument;
		this.element = doc.createElement('div');
		this.element.className = 'mini-toolbar';
		this.element.setAttribute('role', 'toolbar');
		this.element.setAttribute('aria-label', 'Mini toolbar');
		this.element.hidden = true;
		const button = (
			id: string,
			ribbonId: string,
			label: string,
			icon: string,
			action: VisioRibbonAction,
		): RibbonCommand => {
			const spec: CommandSpec = { id: `mini-${id}`, label, icon, size: 'icon', action };
			this.#mirrors.push({ mini: spec.id, ribbon: ribbonId });
			return command(doc, spec);
		};
		this.element.append(
			combo(doc, {
				id: 'mini-font',
				label: 'Font',
				placeholder: 'Font',
				width: 112,
				action: (value) => ({ type: 'font-family', value }),
			}),
			combo(doc, {
				id: 'mini-font-size',
				label: 'Font Size',
				placeholder: 'Size',
				width: 64,
				action: (value) => ({ type: 'font-size', value: Number(value) }),
			}),
			button('grow', 'grow-font', 'Increase Font Size', 'growFont', {
				type: 'font-step',
				direction: 1,
			}),
			button('shrink', 'shrink-font', 'Decrease Font Size', 'shrinkFont', {
				type: 'font-step',
				direction: -1,
			}),
			button('bold', 'bold', 'Bold', 'bold', { type: 'text-toggle', property: 'bold' }),
			button('italic', 'italic', 'Italic', 'italic', { type: 'text-toggle', property: 'italic' }),
			button('underline', 'underline', 'Underline', 'underline', {
				type: 'text-toggle',
				property: 'underline',
			}),
			button('align-left', 'align-left', 'Align Left', 'alignLeft', {
				type: 'text-align',
				axis: 'horizontal',
				value: 'left',
			}),
			button('align-center', 'align-center', 'Center', 'alignCenter', {
				type: 'text-align',
				axis: 'horizontal',
				value: 'center',
			}),
			button('align-right', 'align-right', 'Align Right', 'alignRight', {
				type: 'text-align',
				axis: 'horizontal',
				value: 'right',
			}),
		);
		// The bar is a tool beside the text, not a place to move focus to: a press keeps the editor's.
		this.element.addEventListener('pointerdown', (event) => {
			if (!(event.target instanceof HTMLElement && event.target.closest('office-ui-select')))
				event.preventDefault();
		});
		root.append(this.element);
	}
	/** The Mini Toolbar option: switched off, the bar never shows. */
	set enabled(value: boolean) {
		this.#enabled = value;
		if (!value) this.hide();
	}
	get visible(): boolean {
		return !this.element.hidden;
	}
	/** Copy the state of the ribbon's controls onto the bar's. */
	sync(): void {
		for (const { mini, ribbon } of this.#mirrors) {
			const target = this.element.querySelector<RibbonCommand>(`[command="${mini}"]`);
			const source = this.#ribbon.querySelector<RibbonCommand>(`[command="${ribbon}"]`);
			if (!target || !source) continue;
			target.disabled = source.disabled;
			const pressed = source.getAttribute('pressed');
			if (pressed === null) target.removeAttribute('pressed');
			else target.setAttribute('pressed', pressed);
			target.title = source.title.replace(/:.*$/, '') || target.title;
		}
		for (const [mini, ribbon] of [
			['mini-font', 'font'],
			['mini-font-size', 'font-size'],
		] as const) {
			const target = this.element.querySelector<Combo>(`[data-combo="${mini}"]`);
			const source = this.#ribbon.querySelector<Combo>(`[data-combo="${ribbon}"]`);
			if (!target || !source) continue;
			target.options = source.options.map((option) => ({ ...option }));
			target.value = source.value;
			if (source.hasAttribute('disabled')) target.setAttribute('disabled', '');
			else target.removeAttribute('disabled');
		}
	}
	/**
	 * Show the bar above (`x`, `y`), the pointer or the top of the selection, in viewport pixels;
	 * below it when there is no room above. Nothing shows when the option is off or no command
	 * can act on the selection.
	 */
	showAt(x: number, y: number): void {
		if (!this.#enabled) return;
		this.sync();
		if (!this.element.querySelector('[command]:not([disabled]), office-ui-select:not([disabled])'))
			return;
		this.element.hidden = false;
		const view = this.element.ownerDocument.defaultView;
		const box = this.element.getBoundingClientRect();
		const width = view?.innerWidth ?? 1024;
		const left = Math.min(
			Math.max(MARGIN, x - box.width / 2),
			Math.max(MARGIN, width - box.width - MARGIN),
		);
		const above = y - box.height - GAP;
		this.element.style.left = `${Math.round(left)}px`;
		this.element.style.top = `${Math.round(above >= MARGIN ? above : y + GAP * 3)}px`;
	}
	hide(): void {
		this.element.hidden = true;
	}
	/**
	 * Dismiss on a press elsewhere, Escape, scrolling or zooming the drawing, and when focus leaves
	 * the page. Returns what removes the listeners.
	 */
	wire(viewport: HTMLElement): () => void {
		const doc = this.element.ownerDocument;
		const Abort = doc.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		doc.addEventListener(
			'pointerdown',
			(event) => {
				if (this.visible && !event.composedPath().includes(this.element)) this.hide();
			},
			{ ...options, capture: true },
		);
		doc.addEventListener(
			'keydown',
			(event) => {
				if (event.key === 'Escape') this.hide();
			},
			options,
		);
		viewport.addEventListener('scroll', () => this.hide(), options);
		viewport.addEventListener('wheel', () => this.hide(), { ...options, passive: true });
		doc.defaultView?.addEventListener('blur', () => this.hide(), options);
		return () => events.abort();
	}
}
