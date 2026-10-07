import { LitElement, html, unsafeCSS } from 'lit';
import css from './navigation-drawer.css?raw';
import { modalActiveElement, modalFocusableElements } from '../../dialog/focus';

/** The compact shell's modal navigation. Its slotted content uses the desktop sidebar actions. */
export class TeamsNavigationDrawer extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		open: { type: Boolean, reflect: true },
		heading: { type: String },
	};
	declare open: boolean;
	declare heading: string;
	private compact: MediaQueryList | undefined;
	private readonly resized = () => {
		if (this.open && !this.compact?.matches) this.close();
	};
	constructor() {
		super();
		this.open = false;
		this.heading = 'Teams and channels';
		this.addEventListener('keydown', (event) => this.trapTab(event));
	}
	override connectedCallback(): void {
		super.connectedCallback();
		this.compact = globalThis.matchMedia?.('(max-width: 56.25em)');
		this.compact?.addEventListener('change', this.resized);
	}
	override disconnectedCallback(): void {
		this.compact?.removeEventListener('change', this.resized);
		super.disconnectedCallback();
	}
	protected override updated(): void {
		const dialog = this.renderRoot.querySelector('dialog')!;
		if (this.open && !dialog.open) dialog.showModal();
		else if (!this.open && dialog.open) dialog.close();
	}
	private trapTab(event: KeyboardEvent): void {
		if (!this.open || event.key !== 'Tab') return;
		const close = this.renderRoot.querySelector<HTMLButtonElement>('button')!;
		const stops = [close, ...modalFocusableElements(this)];
		const current = modalActiveElement(this.ownerDocument);
		const index = stops.findIndex((element) => element === current);
		if (index < 0 || (event.shiftKey ? index === 0 : index === stops.length - 1)) {
			event.preventDefault();
			stops[event.shiftKey ? stops.length - 1 : 0]?.focus();
		}
	}
	private close(): void {
		this.dispatchEvent(
			new CustomEvent('teams-navigation-close', { bubbles: true, composed: true }),
		);
	}
	protected override render() {
		return html`<dialog
			aria-label="Workspace navigation"
			@click=${(event: MouseEvent) => {
				const dialog = event.currentTarget as HTMLDialogElement;
				const rect = dialog.getBoundingClientRect();
				if (
					event.target === dialog &&
					(event.clientX < rect.left ||
						event.clientX > rect.right ||
						event.clientY < rect.top ||
						event.clientY > rect.bottom)
				)
					this.close();
			}}
			@cancel=${(event: Event) => {
				event.preventDefault();
				this.close();
			}}
		>
			<header>
				<h2>${this.heading}</h2>
				<button type="button" aria-label="Close navigation" @click=${this.close}>×</button>
			</header>
			<nav aria-label="Workspace channels and calls"><slot></slot></nav>
		</dialog>`;
	}
}

export function defineTeamsNavigationDrawer(
	registry: CustomElementRegistry | undefined = globalThis.customElements,
): void {
	if (registry && !registry.get('teams-navigation-drawer'))
		registry.define('teams-navigation-drawer', TeamsNavigationDrawer);
}
