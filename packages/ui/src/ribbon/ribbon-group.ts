import { html } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base.js';
import { glyph } from '../glyph.js';
import { definer, present } from '../registry.js';
import css from './ribbon-group.css?raw';

type GroupConfig = { launcherEvent: string; collapseEvent: string };

/** A length token resolved on `el`, in pixels (falls back when unset or not in px). */
function tokenPx(el: Element, name: string, fallback: number): number {
	const value = el.ownerDocument.defaultView?.getComputedStyle(el).getPropertyValue(name).trim();
	const px = value?.endsWith('px') ? parseFloat(value) : NaN;
	return Number.isFinite(px) ? px : fallback;
}

/**
 * Labelled group of ribbon commands: `role="group"`, `label` is the caption and name. `launcher`
 * adds Office's corner dialog launcher (`launcher="<command>"` emits `office-command`
 * `{ command }`; `launcher-label` names it; `launcher-disabled` disables it). For narrow windows
 * an overflow controller sets `data-collapsed` (the group becomes one button showing `icon` and
 * the label; activating it emits `office-ribbon-collapse-toggle`) and `data-open` (its commands
 * open in a popup at `--office-ribbon-collapse-x/-y`). The group marks itself `data-compact-row`
 * when every command is one line tall (they centre vertically) and `data-stack` when it holds only
 * small drop-down galleries (`[mode="dropdown"]` without `data-command-large`), which stack in
 * columns. Static `launcherEvent`/`collapseEvent` and `launcherDetail()` let a product subclass
 * keep its published events.
 */
export class OfficeUiRibbonGroup extends OfficeElement {
	static launcherEvent = 'office-command';
	static collapseEvent = 'office-ribbon-collapse-toggle';
	static override styles = controlStyles(css);
	static override properties = {
		label: { type: String },
		// Product hooks read these as attributes, so setting the property writes them through.
		launcher: { type: String, reflect: true },
		launcherLabel: { attribute: 'launcher-label', type: String },
		launcherDisabled: { attribute: 'launcher-disabled', ...flag },
		icon: { type: String, reflect: true },
		open: { attribute: 'data-open', ...flag },
	};
	declare label: string;
	declare launcher: string | null;
	declare launcherLabel: string | null;
	declare launcherDisabled: boolean;
	declare icon: string | null;
	declare open: boolean;
	private observer: ResizeObserver | undefined;

	constructor() {
		super();
		this.label = '';
		this.launcher = null;
		this.launcherLabel = null;
		this.launcherDisabled = false;
		this.icon = null;
		this.open = false;
	}

	/** The launcher's event detail; null emits nothing. */
	protected launcherDetail(): Record<string, unknown> | null {
		const command = this.getAttribute('launcher');
		return command ? { command } : null;
	}

	/** The registered icon on the collapsed face; a product subclass may map names. */
	protected iconName(): string {
		return this.getAttribute('icon') ?? 'grid';
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.observe();
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.observer?.disconnect();
	}

	private observe(): void {
		this.observer?.disconnect();
		// `slotchange` is asynchronous and can fire after removal: observing then would pin the
		// detached subtree (a ResizeObserver holds its targets), so never re-arm.
		if (!this.isConnected) return;
		const Observer = this.ownerDocument.defaultView?.ResizeObserver;
		if (Observer) this.observer ??= new Observer(() => this.measure());
		const slot = this.renderRoot.querySelector('slot');
		for (const el of slot?.assignedElements() ?? []) this.observer?.observe(el);
		this.measure();
	}

	private measure(): void {
		const children = [...this.children];
		const stackable = children.filter((el) => {
			const gallery = el.matches('[mode="dropdown"]') ? el : el.querySelector('[mode="dropdown"]');
			return gallery !== null && !gallery.hasAttribute('data-command-large');
		});
		this.toggleAttribute(
			'data-stack',
			stackable.length > 1 && stackable.length === children.length,
		);
		const tallest = Math.max(0, ...children.map((el) => el.getBoundingClientRect().height));
		if (tallest > 0)
			this.toggleAttribute(
				'data-compact-row',
				tallest <= tokenPx(this, '--office-ribbon-compact-row-max', 36),
			);
	}

	private onLauncher(): void {
		if (present(this.launcherDisabled)) return;
		const detail = this.launcherDetail();
		if (detail) this.fire((this.constructor as unknown as GroupConfig).launcherEvent, detail);
	}

	/** The overflow controller listens for this and owns which popup is open. */
	private onFace(): void {
		this.fire((this.constructor as unknown as GroupConfig).collapseEvent, undefined);
	}

	protected override willUpdate(): void {
		this.setAttribute('role', 'group');
		this.setAttribute('aria-label', this.label);
	}

	protected override render() {
		const launcherLabel = this.launcherLabel ?? `${this.label} options`;
		return html`
			<div class="group">
				<button
					class="face"
					type="button"
					aria-haspopup="true"
					aria-label=${this.label}
					aria-expanded=${String(present(this.open))}
					title=${this.label}
					@click=${this.onFace}
					>${glyph(this.iconName(), 'face-icon')}<span>${this.label}</span>${glyph('chevronDown', 'chev')}</button
				>
				<div class="row">
					<slot @slotchange=${this.observe}></slot>
				</div>
				<div class="foot">
					<span class="caption">${this.label}</span>
					<button
						class="launcher"
						type="button"
						aria-label=${launcherLabel}
						title=${launcherLabel}
						?hidden=${this.launcher === null}
						?disabled=${present(this.launcherDisabled)}
						@click=${this.onLauncher}
						>${glyph('launcher', 'launcher-icon')}</button
					>
				</div>
			</div>
		`;
	}
}

export const defineRibbonGroup = definer('office-ui-ribbon-group', () => OfficeUiRibbonGroup);
