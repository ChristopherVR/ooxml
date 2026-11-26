import { html, nothing } from 'lit';
import { INK_PEN_THICKNESSES_MM, inkThicknessPx, type InkPen } from 'ooxml-core/drawingml';
import { OfficeElement, controlStyles, flag } from '../base';
import { definer, emit } from '../registry';
import css from './ink-pens.css?raw';
import { DEFAULT_INK_PENS_LABELS, INK_PEN_COLORS, type InkPensLabels } from './pen-labels';

export { DEFAULT_INK_PENS_LABELS, INK_PEN_COLORS, type InkPensLabels } from './pen-labels';

const COLOR = /^#[\da-f]{6}$/iu;
const choices = (pen: InkPen) =>
	pen.kind === 'highlighter' ? INK_PEN_THICKNESSES_MM.highlighter : INK_PEN_THICKNESSES_MM.pen;
/** The nearest thickness choice of a pen, in millimetres. */
export function inkPenThicknessMm(pen: InkPen): number {
	return choices(pen).reduce((best, mm) =>
		Math.abs(inkThicknessPx(mm) - pen.width) < Math.abs(inkThicknessPx(best) - pen.width)
			? mm
			: best,
	);
}

/** `office-ink-pen-select`: a pen was picked (or picked again to open its settings). */
export interface InkPenSelectDetail {
	index: number;
}
/** `office-ink-pen-change`: the current pen's colour or thickness (pixels) changed. */
export interface InkPenChangeDetail {
	index: number;
	color?: string;
	width?: number;
}

/**
 * `<office-ui-ink-pens>`: the Draw tab's pen gallery, shared by every product. Each pen is a
 * toggle drawn as a pen in its own colour with a nib as thick as it draws, pressed while it is the
 * pen in use. Picking the pressed pen again opens its settings: thickness choices and colours (a
 * menu kept inside the window), and Delete Pen. Fires `office-ink-pen-select`,
 * `office-ink-pen-change` and `office-ink-pen-remove` (`{ index }`). `disabled` with
 * `disabledReason` turns the strip off and says why.
 */
export class OfficeUiInkPens extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		pens: { attribute: false },
		selected: { type: Number },
		disabled: { ...flag, reflect: true },
		disabledReason: { type: String, attribute: 'disabled-reason' },
		labels: { attribute: false },
	};
	declare pens: readonly InkPen[];
	declare selected: number;
	declare disabled: boolean;
	declare disabledReason: string;
	declare labels: InkPensLabels;
	#open = -1;
	/** The pen that takes the strip's one tab stop. */
	#focus = 0;
	#justOpened = false;
	#listening = false;
	readonly #outside = (event: Event) => {
		if (!event.composedPath().includes(this)) this.close();
	};
	readonly #key = (event: KeyboardEvent) => {
		if (event.key !== 'Escape' || this.#open < 0) return;
		const index = this.#open;
		this.close();
		this.renderRoot.querySelector<HTMLButtonElement>(`.pen[data-index="${index}"]`)?.focus();
	};

	constructor() {
		super();
		this.pens = [];
		this.selected = -1;
		this.disabled = false;
		this.disabledReason = '';
		this.labels = DEFAULT_INK_PENS_LABELS;
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.close();
	}

	/** Closes a pen's settings menu. */
	close(): void {
		if (this.#open < 0) return;
		this.#open = -1;
		this.#listen(false);
		this.requestUpdate();
	}

	#listen(on: boolean): void {
		if (on === this.#listening) return;
		this.#listening = on;
		const doc = this.ownerDocument;
		if (on) {
			doc.addEventListener('pointerdown', this.#outside, true);
			doc.addEventListener('keydown', this.#key, true);
		} else {
			doc.removeEventListener('pointerdown', this.#outside, true);
			doc.removeEventListener('keydown', this.#key, true);
		}
	}

	#name(pen: InkPen): string {
		const labels = this.labels;
		const color =
			labels.colorNames?.[
				INK_PEN_COLORS.find((c) => c.value === pen.color.toUpperCase())?.name ?? ''
			];
		return labels.pen
			.replace('{kind}', labels.kinds[pen.kind] ?? pen.kind)
			.replace('{size}', String(inkPenThicknessMm(pen)))
			.replace('{color}', color ?? pen.color.toUpperCase());
	}

	#pick(index: number): void {
		if (this.disabled) return;
		if (index === this.selected) {
			this.#open = this.#open === index ? -1 : index;
			this.#justOpened = this.#open >= 0;
			this.#listen(this.#open >= 0);
			this.requestUpdate();
		} else this.close();
		emit<InkPenSelectDetail>(this, 'office-ink-pen-select', { index });
	}

	#change(change: Omit<InkPenChangeDetail, 'index'>): void {
		const index = this.#open;
		if (index < 0) return;
		emit<InkPenChangeDetail>(this, 'office-ink-pen-change', { index, ...change });
	}

	protected override updated(): void {
		const menu = this.renderRoot.querySelector<HTMLElement>('.menu');
		if (menu && this.#justOpened) {
			this.#justOpened = false;
			queueMicrotask(() =>
				menu.querySelector<HTMLElement>('[aria-checked="true"], button')?.focus(),
			);
		}
		const anchor = this.renderRoot.querySelector<HTMLElement>(`.pen[data-index="${this.#open}"]`);
		if (!menu || !anchor || this.#open < 0) return;
		// Kept inside the window, under the pen when it fits and above it otherwise.
		const view = this.ownerDocument.defaultView;
		const width = view?.innerWidth ?? 1024;
		const height = view?.innerHeight ?? 768;
		const box = anchor.getBoundingClientRect();
		const own = menu.getBoundingClientRect();
		const left = Math.max(8, Math.min(box.left, width - own.width - 8));
		const below = box.bottom + 4;
		const top = below + own.height <= height - 8 ? below : Math.max(8, box.top - own.height - 4);
		menu.style.left = `${left}px`;
		menu.style.top = `${top}px`;
	}

	/** A pen drawn as Office draws it in the gallery: a stroke in its colour and thickness on paper. */
	#swatch(pen: InkPen) {
		const color = COLOR.test(pen.color) ? pen.color : '#000000';
		const width = Math.max(1.5, Math.min(10, pen.width));
		const opacity = pen.kind === 'highlighter' ? 0.5 : pen.kind === 'pencil' ? 0.75 : 1;
		return html`<svg viewBox="0 0 22 40" aria-hidden="true">
			<rect class="paper" x="0.5" y="0.5" width="21" height="39" rx="3"></rect>
			<path
				d="M11 6C3 12 19 18 11 24S5 32 11 35"
				fill="none"
				stroke=${color}
				stroke-width=${width}
				stroke-opacity=${opacity}
				stroke-linecap=${pen.kind === 'highlighter' ? 'butt' : 'round'}
			></path>
		</svg>`;
	}

	#keydown(event: KeyboardEvent): void {
		const buttons = [...this.renderRoot.querySelectorAll<HTMLButtonElement>('.pen')];
		const at = buttons.indexOf(event.target as HTMLButtonElement);
		if (at < 0) return;
		const step = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0;
		const next = event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 : at + step;
		// The strip is one stop; arrows move inside it and leave it at its ends.
		if (
			(!step && event.key !== 'Home' && event.key !== 'End') ||
			next < 0 ||
			next >= buttons.length
		)
			return;
		event.preventDefault();
		event.stopPropagation();
		this.#focus = next;
		this.requestUpdate();
		buttons[next]?.focus();
	}

	protected override render() {
		const labels = this.labels;
		const open = this.pens[this.#open];
		const reason = this.disabled ? this.disabledReason : '';
		return html`<div class="strip" role="group" aria-label=${labels.group}
				>${this.pens.map((pen, index) => {
					const name = this.#name(pen);
					const stop =
						this.selected >= 0 ? this.selected : Math.min(this.#focus, this.pens.length - 1);
					return html`<button
						type="button"
						class="pen"
						data-index=${index}
						tabindex=${index === stop ? '0' : '-1'}
						@keydown=${(event: KeyboardEvent) => this.#keydown(event)}
						aria-label=${name}
						title=${reason ? `${name}: ${reason}` : `${name}\n${labels.description}`}
						aria-pressed=${String(index === this.selected)}
						aria-haspopup=${index === this.selected ? 'true' : 'false'}
						aria-expanded=${index === this.selected ? String(this.#open === index) : nothing}
						aria-description=${reason || nothing}
						?disabled=${this.disabled}
						@click=${() => this.#pick(index)}
						>${this.#swatch(pen)}${index === this.selected ? html`<svg class="caret" viewBox="0 0 10 6" aria-hidden="true"><path d="M1 1l4 4 4-4"></path></svg>` : nothing}</button
					>`;
				})}</div
			>${
				open
					? html`<div class="menu" role="dialog" aria-label=${this.#name(open)}
							><div class="heading" id="thickness">${labels.thickness}</div
							><div class="thicknesses" role="radiogroup" aria-labelledby="thickness"
								>${choices(open).map((mm) => {
									const px = inkThicknessPx(mm);
									const on = inkPenThicknessMm(open) === mm;
									return html`<button
										type="button"
										class="thickness"
										role="radio"
										aria-checked=${String(on)}
										aria-label=${`${mm} mm`}
										title=${`${mm} mm`}
										@click=${() => this.#change({ width: px })}
										><span
											style=${`height:${Math.max(1, Math.min(14, px))}px;color:${open.color}`}
										></span
									></button>`;
								})}</div
							><div class="heading" id="colors">${labels.colors}</div
							><div class="colors" role="radiogroup" aria-labelledby="colors"
								>${INK_PEN_COLORS.map(({ value, name }) => {
									const label = labels.colorNames?.[name] ?? name;
									return html`<button
										type="button"
										class="swatch"
										role="radio"
										aria-checked=${String(open.color.toUpperCase() === value)}
										aria-label=${label}
										title=${label}
										@click=${() => this.#change({ color: value })}
										><span style=${`background:${value}`}></span
									></button>`;
								})}</div
							><label class="more"
								>${labels.moreColors}<input
									type="color"
									.value=${open.color.toLowerCase()}
									@change=${(event: Event) => this.#change({ color: (event.target as HTMLInputElement).value.toUpperCase() })} /></label
							><button
								type="button"
								class="remove"
								?disabled=${this.pens.length <= 1}
								@click=${() => {
									const index = this.#open;
									this.close();
									emit(this, 'office-ink-pen-remove', { index });
								}}
								>${labels.remove}</button
							></div
						>`
					: nothing
			}`;
	}
}

export const defineInkPens = definer('office-ui-ink-pens', () => OfficeUiInkPens);
