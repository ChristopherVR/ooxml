import { html, nothing } from 'lit';
import type { LocalAddIn } from 'ooxml-core/automation/add-ins';
import { OfficeElement, controlStyles } from '../base';
import { glyph } from '../glyph';
import { definer } from '../registry';
import { addInSettings, type OpenAddInsOptions } from '../ribbon/add-in-dialog';
import { editorAddIns } from '../ribbon/add-in-installed-tabs';
import {
	addInError,
	addInIssueLabels,
	installAddIn,
	localAddIns,
	readAddInFile,
} from '../ribbon/add-in-local-dialog';
import { renderAddInConsent } from './add-ins-consent';
import { localAddInText } from '../ribbon/add-in-local-strings';
import { defineDialog, type OfficeUiDialog } from './dialog';
import css from './add-ins-dialog.css?raw';

const PIVOTS = ['My Add-ins', 'Admin Managed', 'Store'] as const;
type Pivot = (typeof PIVOTS)[number];

const MARKETPLACE = 'https://marketplace.microsoft.com/en-us/marketplace/apps?product=office';

function handoffUrl(link: string): string | undefined {
	try {
		const url = new URL(link);
		if (url.protocol === 'https:' && !url.username && !url.password) return url.href;
	} catch {
		// An invalid host configuration has no handoff.
	}
}

function addInSession(options: OpenAddInsOptions, restoreFocus: () => void) {
	const { host, locale, t, product } = options;
	return {
		...addInSettings(host),
		installed: localAddIns(host),
		tabs: editorAddIns(host),
		t: localAddInText(locale, t),
		product,
		locale,
		restoreFocus,
	};
}

type AddInSession = ReturnType<typeof addInSession>;
type AddInPreview = { addIn: LocalAddIn; previous: LocalAddIn | undefined };

/** On-demand Office Add-ins browser with local manifest review and a host-configured storefront. */
export class OfficeUiAddInsDialog extends OfficeElement {
	static override styles = controlStyles(css);

	#session?: AddInSession;
	#pivot: Pivot = 'My Add-ins';
	#query = '';
	#message = '';
	#preview: AddInPreview | undefined;
	#pending: AbortController | undefined;
	#afterClose?: () => void;

	show(options: OpenAddInsOptions, restoreFocus: () => void): void {
		this.#session = addInSession(options, restoreFocus);
		if (options.mode === 'getAddIns') {
			this.#pivot = this.#session.storefront === 'off' ? 'Admin Managed' : 'Store';
		}
		this.requestUpdate();
		this.shell?.show();
		this.renderRoot.querySelector<HTMLInputElement>('input[type="search"]')?.focus();
	}

	private get shell(): OfficeUiDialog | null {
		return this.renderRoot.querySelector<OfficeUiDialog>('office-ui-dialog');
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.#pending?.abort();
	}

	private close(): void {
		queueMicrotask(() => {
			if (this.shell?.open) return;
			this.remove();
			if (this.#afterClose) this.#afterClose();
			else this.#session?.restoreFocus();
		});
	}

	private pivot(pivot: Pivot): void {
		this.#pivot = pivot;
		this.#message = '';
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

	private async read(source: (signal: AbortSignal) => File | Promise<File>): Promise<void> {
		const session = this.#session;
		if (!session || session.product === 'visio' || this.#pending) return;
		const pending = (this.#pending = new AbortController());
		this.#message = session.t('Loading add-ins…');
		this.requestUpdate();
		try {
			const file = await source(pending.signal);
			if (pending.signal.aborted || !this.shell?.open) return;
			const addIn = await readAddInFile(file, session.product, session.locale);
			if (pending.signal.aborted || !this.shell?.open) return;
			this.#preview = { addIn, previous: session.installed.get(addIn.manifest.id) };
			this.#message = '';
		} catch (error) {
			this.#message = session.t(addInError(error));
		} finally {
			this.#pending = undefined;
			if (this.shell?.open) {
				this.requestUpdate();
				const focus = this.#preview ? '#consent-title' : '#upload';
				this.renderRoot.querySelector<HTMLElement>(focus)?.focus();
			}
		}
	}

	private confirm(): void {
		const preview = this.#preview;
		const session = this.#session;
		if (!preview || !session) return;
		if (installAddIn(session.installed, preview.addIn, preview.previous)) this.shell?.close();
		else {
			this.#message = session.t('Could not update the add-in.');
			this.requestUpdate();
		}
	}

	private cancel(): void {
		this.#preview = undefined;
		this.#message = '';
		this.requestUpdate();
		this.renderRoot.querySelector<HTMLInputElement>('input[type="search"]')?.focus();
	}

	private async change(run: () => void | Promise<void>): Promise<void> {
		if (this.#pending) return;
		this.#pending = new AbortController();
		this.#message = '';
		this.requestUpdate();
		try {
			await run();
			this.shell?.close();
		} catch {
			this.#message = this.#session?.t('Could not update the add-in.') ?? '';
		} finally {
			this.#pending = undefined;
			if (this.shell?.open) this.requestUpdate();
		}
	}

	private tile(label: string, actions: unknown, details?: unknown) {
		return html`<li class="tile"
			>${glyph('addIns', 'tile-icon')}<div class="tile-body">
				<p class="tile-name">${label}</p>${details}
				<div class="tile-actions">${actions}</div>
			</div></li
		>`;
	}

	private list(session: AddInSession) {
		const { installed, tabs, catalog, storefront, t, locale, product } = session;
		const query = this.#query.trim().toLocaleLowerCase(locale);
		const matches = (label: string) => label.toLocaleLowerCase(locale).includes(query);
		const action = (label: string, name: string, run: () => void, disabled = false) =>
			html`<button
				type="button"
				aria-label=${`${t(label)} ${name}`}
				?disabled=${!!this.#pending || disabled}
				@click=${run}
				>${t(label)}</button
			>`;
		const installedIds = new Set(tabs.map(({ id }) => id));
		const offers =
			this.#pivot === 'Store' && storefront !== 'off' && storefront.kind === 'catalog'
				? storefront.entries
				: [];
		const tiles =
			this.#pivot === 'My Add-ins'
				? [
						...tabs
							.filter(({ label }) => matches(label))
							.map(({ id, label, open }) => {
								const remove = catalog.find((entry) => entry.id === id)?.remove;
								return this.tile(label, [
									action('Open', label, () => {
										this.#afterClose = open;
										this.shell?.close();
									}),
									remove ? action('Remove', label, () => void this.change(remove)) : nothing,
								]);
							}),
						...[...installed.values()]
							.filter(({ manifest }) => matches(manifest.name))
							.map(({ manifest }) =>
								this.tile(
									`${manifest.name} (${manifest.version})`,
									action(
										'Remove',
										manifest.name,
										() =>
											void this.change(() => {
												installed.delete(manifest.id);
											}),
									),
									html`<p>${t('Installed for this editor. Running add-ins is unavailable.')}</p>
										${addInIssueLabels(manifest, t).map((line) => html`<p>${line}</p>`)}`,
								),
							),
					]
				: this.#pivot === 'Admin Managed'
					? catalog
							.filter(({ id, label }) => !installedIds.has(id) && matches(label))
							.map(({ label, install }) =>
								this.tile(
									label,
									action('Install', label, () => void this.change(install)),
								),
							)
					: offers
							.filter(({ label }) => matches(label))
							.map(({ label, read }) =>
								this.tile(
									label,
									action('Install', label, () => void this.read(read), product === 'visio'),
								),
							);
		const handoff =
			this.#pivot === 'Store' && storefront !== 'off' && storefront.kind === 'microsoft'
				? handoffUrl(storefront.url ?? MARKETPLACE)
				: undefined;
		if (handoff && storefront !== 'off' && storefront.kind === 'microsoft') {
			return html`<p
					>${t('Browse add-ins for Microsoft Office. This editor runs add-ins supplied by its host.')}</p
				>
				<a href=${handoff} target="_blank" rel="noopener noreferrer"
					>${storefront.label ?? t('Browse Microsoft Marketplace')}</a
				>`;
		}
		return tiles.length
			? html`<ul class="tiles"
					>${tiles}</ul
				>`
			: html`<p class="empty"
					>${t(query ? 'No matching add-ins.' : 'No add-ins are available in this editor.')}</p
				>`;
	}

	protected override render() {
		const session = this.#session;
		if (!session) return nothing;
		const { t, storefront, product } = session;
		const pivots = storefront === 'off' ? PIVOTS.slice(0, 2) : PIVOTS;
		const unsupported = product === 'visio';
		const reason = t('Local add-in installation is unavailable in Visio.');
		const search = t('Search add-ins');
		return html`<office-ui-dialog
			.heading=${t('Office Add-ins')}
			.closeLabel=${t('Close dialog')}
			@office-dialog-close=${() => this.close()}
		>
			${
				this.#preview
					? renderAddInConsent(this.#preview.addIn, this.#preview.previous, t, {
							cancel: () => this.cancel(),
							confirm: () => this.confirm(),
						})
					: html` <div
								class="pivots"
								role="tablist"
								aria-label=${t('Office Add-ins')}
								@keydown=${(event: KeyboardEvent) => this.pivotKey(event, pivots)}
							>
								${pivots.map(
									(pivot, index) =>
										html`<button
											type="button"
											role="tab"
											id=${`pivot-${index}`}
											aria-selected=${String(this.#pivot === pivot)}
											aria-controls="list"
											tabindex=${this.#pivot === pivot ? 0 : -1}
											@click=${() => this.pivot(pivot)}
											>${t(pivot)}</button
										>`,
								)}
							</div>
							<div
								id="list"
								role="tabpanel"
								aria-labelledby=${`pivot-${pivots.indexOf(this.#pivot)}`}
							>
								<div class="tools">
									<input
										type="search"
										aria-label=${search}
										placeholder=${search}
										maxlength="120"
										.value=${this.#query}
										@input=${(event: InputEvent) => {
											if (!(event.currentTarget instanceof HTMLInputElement)) return;
											this.#query = event.currentTarget.value;
											this.requestUpdate();
										}}
									/>
									<button
										id="upload"
										class="link"
										type="button"
										?disabled=${unsupported || !!this.#pending}
										title=${unsupported ? reason : t('Choose a local add-in file')}
										@click=${() => this.renderRoot.querySelector<HTMLInputElement>('input[type="file"]')?.click()}
										>${t('Upload My Add-in')}</button
									>
									<input
										hidden
										type="file"
										accept=".xml,.json,.zip"
										aria-label=${t('Choose a local add-in file')}
										?disabled=${unsupported || !!this.#pending}
										@change=${this.upload}
									/>
								</div>
								${unsupported ? html`<p class="empty">${reason}</p>` : nothing}
								${this.list(session)}
							</div>`
			}
			<p class="status" role="status">${this.#message}</p>
			<button type="button" slot="footer" @click=${() => this.shell?.close()}>${t('Close')}</button>
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
