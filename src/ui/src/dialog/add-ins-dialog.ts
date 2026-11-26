import { html, nothing } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import { mountedAddInRestore, type OpenAddInsOptions } from '../ribbon/add-in-dialog';
import { addInInstallFeedback } from '../ribbon/add-in-install-feedback';
import {
	ADD_IN_RUNTIME_CHANGED,
	ADD_IN_RUNTIME_REFRESHED,
} from '../ribbon/add-in-runtime-settings';
import { marketplaceAssetId } from '../ribbon/add-in-marketplace';
import { renderAddInConsent, renderAddInConsentActions } from './add-ins-consent';
import { renderAddInBrowser, type Pivot } from './add-ins-browser';
import { addInSession, type AddInSession, type AddInPreview } from './add-ins-session';
import { defineDialog, type OfficeUiDialog } from './dialog';
import css from './add-ins-dialog.css?raw';

/** On-demand Office Add-ins browser with local manifest review and a host-configured storefront. */
export class OfficeUiAddInsDialog extends OfficeElement {
	static override styles = controlStyles(css);

	#session?: AddInSession;
	#pivot: Pivot = 'My Add-ins';
	#query = '';
	#message = '';
	#error = '';
	#preview: AddInPreview | undefined;
	#pending: AbortController | undefined;
	#afterClose?: () => void;
	#settingsEvents = new AbortController();

	show(options: OpenAddInsOptions, restoreFocus: () => void): void {
		this.#session = addInSession(options, restoreFocus);
		this.#session.syncTabs();
		options.host.addEventListener(ADD_IN_RUNTIME_REFRESHED, () => this.requestUpdate(), {
			signal: this.#settingsEvents.signal,
		});
		for (const scope of options.scopes)
			scope.addEventListener(ADD_IN_RUNTIME_CHANGED, () => this.requestUpdate(), {
				signal: this.#settingsEvents.signal,
			});
		this.#pivot =
			options.mode === 'getAddIns' && this.#session.storefront !== 'off' ? 'Store' : 'My Add-ins';
		this.requestUpdate();
		this.shell?.show();
		this.renderRoot.querySelector<HTMLInputElement>('input[type="search"]')?.focus();
		const mounted = mountedAddInRestore(options.host);
		if (this.#session.restore)
			void this.restore(this.#session, mounted ? () => mounted : this.#session.restore);
	}

	private async restore(session: AddInSession, restore: NonNullable<AddInSession['restore']>) {
		const pending = (this.#pending = new AbortController());
		this.requestUpdate();
		try {
			await restore(pending.signal);
			if (!pending.signal.aborted) session.syncTabs();
		} catch {
			this.#error = session.t('Could not load saved add-ins.');
		} finally {
			this.#pending = undefined;
			if (!pending.signal.aborted) this.requestUpdate();
		}
	}

	private get shell(): OfficeUiDialog | null {
		return this.renderRoot.querySelector<OfficeUiDialog>('office-ui-dialog');
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.#pending?.abort();
		this.#settingsEvents.abort();
	}

	private close(): void {
		this.remove();
		if (this.#afterClose) this.#afterClose();
		else this.#session?.restoreFocus();
	}

	private pivot(pivot: Pivot): void {
		this.#pivot = pivot;
		this.#message = '';
		this.#error = '';
		this.requestUpdate();
	}

	private pivotKey(event: KeyboardEvent, pivots: readonly Pivot[]): void {
		const index = pivots.indexOf(this.#pivot);
		const target =
			event.key === 'Home'
				? pivots[0]
				: event.key === 'End'
					? pivots.at(-1)
					: event.key === 'ArrowRight'
						? pivots[(index + 1) % pivots.length]
						: event.key === 'ArrowLeft'
							? pivots[(index + pivots.length - 1) % pivots.length]
							: undefined;
		if (!target) return;
		event.preventDefault();
		this.pivot(target);
		this.renderRoot.querySelector<HTMLButtonElement>('[aria-selected="true"]')?.focus();
	}

	private upload = (event: Event): void => {
		const input = event.currentTarget;
		if (!(input instanceof HTMLInputElement)) return;
		const file = input.files?.[0];
		input.value = '';
		if (file) void this.read(() => file);
	};

	private resolve = (event: SubmitEvent): void => {
		event.preventDefault();
		const session = this.#session;
		if (!session?.resolver || this.#pending) return;
		const input = this.renderRoot.querySelector<HTMLInputElement>('#marketplace-link');
		const assetId = marketplaceAssetId(input?.value ?? '');
		if (!assetId) {
			this.showError(session.t('Enter a Microsoft Marketplace link or asset ID.'));
			input?.focus();
			return;
		}
		const resolver = session.resolver;
		void this.read((signal) => resolver(assetId, signal), '#marketplace-link');
	};

	private async read(
		source: (signal: AbortSignal) => File | Promise<File | Uint8Array | ArrayBuffer>,
		returnFocus = '#upload',
	): Promise<void> {
		const session = this.#session;
		if (!session || session.product === 'visio' || this.#pending) return;
		const pending = (this.#pending = new AbortController());
		this.#error = '';
		this.#message = session.t('Loading add-ins…');
		this.requestUpdate();
		let manifestReader: typeof import('../ribbon/add-in-manifest-file') | undefined;
		try {
			manifestReader = await import('../ribbon/add-in-manifest-file');
			if (pending.signal.aborted || !this.shell?.open) return;
			const file = await source(pending.signal);
			if (pending.signal.aborted || !this.shell?.open) return;
			const { addIn, bytes } = await manifestReader.readAddInInput(file, {
				...session.manifestOptions(),
				host: session.product,
			});
			if (pending.signal.aborted || !this.shell?.open) return;
			this.#preview = { addIn, bytes, previous: session.installed.get(addIn.manifest.id) };
			this.#message = '';
			this.#error = '';
		} catch (error) {
			this.#message = '';
			this.showError(
				session.t(manifestReader?.addInError(error) ?? 'Could not update the add-in.'),
			);
		} finally {
			this.#pending = undefined;
			if (this.shell?.open) {
				this.requestUpdate();
				const focus = this.#preview ? '#consent-title' : returnFocus;
				this.renderRoot.querySelector<HTMLElement>(focus)?.focus();
			}
		}
	}

	private renderMarketplace(session: AddInSession) {
		if (this.#pivot !== 'My Add-ins' || !session.resolver || session.product === 'visio')
			return nothing;
		const label = session.t('Add from Marketplace link');
		return html`<form class="tools" @submit=${this.resolve}>
			<label for="marketplace-link">${label}</label>
			<input id="marketplace-link" type="text" maxlength="2048" ?disabled=${!!this.#pending} />
			<button type="submit" ?disabled=${!!this.#pending}>${session.t('Review add-in')}</button>
		</form>`;
	}

	private confirm(): void {
		const preview = this.#preview;
		const session = this.#session;
		if (!preview || !session) return;
		void this.change(async () => {
			await session.install(preview);
			this.#afterClose = () => {
				if (!addInInstallFeedback(session.host, preview.addIn, session.t)) session.restoreFocus();
			};
		});
	}

	private cancel(): void {
		this.#preview = undefined;
		this.#message = '';
		this.#error = '';
		this.requestUpdate();
		this.renderRoot.querySelector<HTMLInputElement>('input[type="search"]')?.focus();
	}

	private async change(run: () => void | Promise<void>): Promise<void> {
		if (this.#pending) return;
		this.#pending = new AbortController();
		this.#message = '';
		this.#error = '';
		this.requestUpdate();
		try {
			await run();
			this.shell?.close();
		} catch {
			this.showError(this.#session?.t('Could not update the add-in.') ?? '');
		} finally {
			this.#pending = undefined;
			if (this.shell?.open) this.requestUpdate();
		}
	}

	private showError(message: string): void {
		this.#error = message;
		this.requestUpdate();
		void this.updateComplete.then(() =>
			this.renderRoot.querySelector('[role="alert"]')?.scrollIntoView?.({ block: 'nearest' }),
		);
	}

	private renderError() {
		return html`<p id="add-in-error" class="error" role="alert">${this.#error}</p>`;
	}

	protected override render() {
		const session = this.#session;
		if (!session) return nothing;
		const { t } = session;
		return html`<office-ui-dialog
			.heading=${t('Office Add-ins')}
			.closeLabel=${t('Close dialog')}
			@office-dialog-closed=${(event: Event) => {
				if (event.target === event.currentTarget) this.close();
			}}
		>
			${
				this.#preview
					? html`${this.renderError()}${renderAddInConsent(
							this.#preview.addIn,
							this.#preview.previous,
							t,
							{
								persistent: Boolean(session.restore),
								reason: session.launcher.pane(this.#preview.addIn).reason,
							},
						)}`
					: renderAddInBrowser({
							session,
							pivot: this.#pivot,
							query: this.#query,
							busy: Boolean(this.#pending),
							error: this.#error,
							alert: this.renderError(),
							marketplace: this.renderMarketplace(session),
							select: (pivot) => this.pivot(pivot),
							key: (event, pivots) => this.pivotKey(event, pivots),
							search: (event) => {
								if (!(event.currentTarget instanceof HTMLInputElement)) return;
								this.#query = event.currentTarget.value;
								this.requestUpdate();
							},
							pick: () =>
								this.renderRoot.querySelector<HTMLInputElement>('input[type="file"]')?.click(),
							upload: this.upload,
							open: (tab) => {
								this.#afterClose = tab;
								this.shell?.close();
							},
							change: (change) => void this.change(change),
							read: (read) => void this.read(read),
						})
			}
			<p class="status" role="status">${this.#message}</p>
			${
				this.#preview
					? renderAddInConsentActions(this.#preview.addIn, this.#preview.previous, t, {
							cancel: () => this.cancel(),
							confirm: () => this.confirm(),
							busy: Boolean(this.#pending),
						})
					: html`<button type="button" slot="footer" @click=${() => this.shell?.close()}
							>${t('Close')}</button
						>`
			}
		</office-ui-dialog>`;
	}
}

export const defineAddInsDialog = definer('office-ui-add-ins-dialog', () => OfficeUiAddInsDialog, [
	defineDialog,
]);

export function mountAddInsDialog(options: OpenAddInsOptions, restoreFocus: () => void): void {
	const { host } = options;
	const registry = host.ownerDocument.defaultView?.customElements;
	defineAddInsDialog(registry);
	const dialog = host.ownerDocument.createElement('office-ui-add-ins-dialog');
	// Independently bundled UI entries can share the registered constructor.
	if (!('show' in dialog) || typeof dialog.show !== 'function') return;
	host.append(dialog);
	dialog.show(options, restoreFocus);
}
