import { html, nothing } from 'lit';
import type { CloudFileProvider } from 'ooxml-core/embedding';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import css from './connected-services.css?raw';

/** Every user-visible string of `<office-ui-connected-services>`; set `labels` (partial). */
export interface OfficeConnectedServicesLabels {
	heading: string;
	/** Shown under the heading. */
	description: string;
	/** A provider with no account; `{provider}` is its label. */
	notConnected: string;
	connect: string;
	disconnect: string;
	/** A connected provider whose host gives no sign-out (the host manages the session). */
	managed: string;
	/** A failed connect or disconnect; `{provider}` is its label. */
	failed: string;
}

export const DEFAULT_CONNECTED_SERVICES_LABELS: OfficeConnectedServicesLabels = {
	heading: 'Connected services',
	description: 'Link your storage to open and save files from every app.',
	notConnected: 'Not connected',
	connect: 'Connect',
	disconnect: 'Disconnect',
	managed: 'Managed by your organization',
	failed: 'Could not update {provider}. Try again.',
};

export type OfficeConnectedServiceChangeEvent = CustomEvent<{
	providerId: string;
	connected: boolean;
}>;

/**
 * `<office-ui-connected-services>`: the Account page's list of linked storage, the same in every
 * product and the Office app. One row per `CloudFileProvider` with its account (name and email)
 * and Connect or Disconnect, which call the provider's own `signIn` and `signOut`: the host's
 * sign-in, never one of the element's. A provider without `signOut` shows that its session is
 * managed by the host. It follows `subscribe`, so signing in from File > Open updates it too.
 * Properties: `providers`, `labels` (see `DEFAULT_CONNECTED_SERVICES_LABELS`), `heading-level`.
 * Event: `office-connected-service-change` `{ providerId, connected }` after a change made here.
 * Without providers it renders nothing.
 */
export class OfficeUiConnectedServices extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		providers: { attribute: false },
		labels: { attribute: false },
		headingLevel: { type: Number, attribute: 'heading-level' },
		busy: { state: true },
		failed: { state: true },
	};
	declare providers: readonly CloudFileProvider[];
	declare labels: Partial<OfficeConnectedServicesLabels>;
	declare headingLevel: number;
	declare busy: string | null;
	declare failed: string | null;
	private subscriptions: (() => void)[] = [];

	constructor() {
		super();
		this.providers = [];
		this.labels = {};
		this.headingLevel = 2;
		this.busy = null;
		this.failed = null;
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.subscribe();
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.unsubscribe();
	}

	protected override willUpdate(changed: Map<PropertyKey, unknown>): void {
		if (changed.has('providers') && this.isConnected) this.subscribe();
	}

	private label(key: keyof OfficeConnectedServicesLabels, provider?: CloudFileProvider): string {
		const text = this.labels[key] ?? DEFAULT_CONNECTED_SERVICES_LABELS[key];
		return provider ? text.replaceAll('{provider}', provider.label) : text;
	}

	private subscribe(): void {
		this.unsubscribe();
		for (const provider of this.providers)
			this.subscriptions.push(provider.subscribe?.(() => this.requestUpdate()) ?? (() => {}));
	}

	private unsubscribe(): void {
		for (const stop of this.subscriptions.splice(0)) stop();
	}

	private async toggle(provider: CloudFileProvider): Promise<void> {
		const connect = !provider.account();
		this.busy = provider.id;
		this.failed = null;
		try {
			if (connect) await provider.signIn();
			else await provider.signOut?.();
			this.fire('office-connected-service-change', {
				providerId: provider.id,
				connected: Boolean(provider.account()),
			});
		} catch {
			this.failed = provider.id;
		} finally {
			this.busy = null;
			this.requestUpdate();
		}
	}

	private row(provider: CloudFileProvider) {
		const account = provider.account();
		const busy = this.busy === provider.id;
		const action = account
			? provider.signOut
				? this.label('disconnect')
				: null
			: this.label('connect');
		return html`<li class="service" data-provider=${provider.id}>
			<span class="badge" aria-hidden="true">${provider.label.slice(0, 1)}</span>
			<span class="details">
				<span class="name">${provider.label}</span>
				<span class="account"
					>${
						account
							? `${account.name}${account.email && account.email !== account.name ? ` (${account.email})` : ''}`
							: this.label('notConnected')
					}</span
				>
				${
					this.failed === provider.id
						? html`<span class="error" role="alert">${this.label('failed', provider)}</span>`
						: nothing
				}
			</span>
			${
				action
					? html`<button
							type="button"
							class=${account ? 'secondary' : 'primary'}
							?disabled=${busy}
							aria-busy=${busy ? 'true' : 'false'}
							aria-label=${`${action} ${provider.label}`}
							@click=${() => void this.toggle(provider)}
						>
							${action}
						</button>`
					: html`<span class="managed">${this.label('managed')}</span>`
			}
		</li>`;
	}

	protected override render() {
		if (!this.providers.length) return nothing;
		const level = Math.min(6, Math.max(1, Math.round(this.headingLevel) || 2));
		return html`<section aria-labelledby="heading">
			<div id="heading" class="heading" role="heading" aria-level=${level}
				>${this.label('heading')}</div
			>
			<p class="description">${this.label('description')}</p>
			<ul class="services">
				${this.providers.map((provider) => this.row(provider))}
			</ul>
		</section>`;
	}
}

export const defineConnectedServices = definer(
	'office-ui-connected-services',
	() => OfficeUiConnectedServices,
);
