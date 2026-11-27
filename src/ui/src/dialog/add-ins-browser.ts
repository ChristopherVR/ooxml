import { html, nothing, type TemplateResult } from 'lit';
import type { OfficeAddInOffer } from '../ribbon/add-in-settings';
import { LOCAL_INSTALL_REASON_ID, renderAddInList } from './add-ins-list';
import type { AddInSession } from './add-ins-session';

export const ADD_IN_PIVOTS = ['My Add-ins', 'Admin Managed', 'Store'] as const;
export type Pivot = (typeof ADD_IN_PIVOTS)[number];

interface AddInBrowserOptions {
	session: AddInSession;
	pivot: Pivot;
	query: string;
	busy: boolean;
	error: string;
	alert: TemplateResult;
	marketplace: unknown;
	select: (pivot: Pivot) => void;
	key: (event: KeyboardEvent, pivots: readonly Pivot[]) => void;
	search: (event: InputEvent) => void;
	pick: () => void;
	upload: (event: Event) => void;
	open: (run: () => void) => void;
	change: (run: () => void | Promise<void>) => void;
	read: (run: OfficeAddInOffer['read']) => void;
}

export function renderAddInBrowser(options: AddInBrowserOptions) {
	const { session } = options;
	const { t, storefront, product } = session;
	const pivots = storefront === 'off' ? ADD_IN_PIVOTS.slice(0, 2) : ADD_IN_PIVOTS;
	const unsupported = product === 'visio';
	const reason = t('Local add-in installation is unavailable in Visio.');
	const search = t('Search add-ins');
	return html`<div
			class="pivots"
			role="tablist"
			aria-label=${t('Office Add-ins')}
			@keydown=${(event: KeyboardEvent) => options.key(event, pivots)}
		>
			${pivots.map(
				(pivot, index) =>
					html`<button
						type="button"
						role="tab"
						id=${`pivot-${index}`}
						aria-selected=${String(options.pivot === pivot)}
						aria-controls="list"
						tabindex=${options.pivot === pivot ? 0 : -1}
						@click=${() => options.select(pivot)}
						>${t(pivot)}</button
					>`,
			)}
		</div>
		<div id="list" role="tabpanel" aria-labelledby=${`pivot-${pivots.indexOf(options.pivot)}`}>
			<div class="tools">
				<input
					type="search"
					aria-label=${search}
					placeholder=${search}
					maxlength="120"
					.value=${options.query}
					@input=${options.search}
				/>
				<button
					id="upload"
					class="link"
					type="button"
					?disabled=${unsupported || options.busy}
					aria-describedby=${unsupported ? LOCAL_INSTALL_REASON_ID : options.error ? 'add-in-error' : nothing}
					@click=${options.pick}
					>${t('Upload My Add-in')}</button
				>
				<input
					hidden
					type="file"
					accept=".xml,.json,.zip"
					aria-label=${t('Choose a local add-in file')}
					?disabled=${unsupported || options.busy}
					@change=${options.upload}
				/>
			</div>
			${options.alert}
			${unsupported ? html`<p class="empty" id=${LOCAL_INSTALL_REASON_ID}>${reason}</p>` : nothing}
			${options.marketplace}
			${renderAddInList({
				session,
				pivot: options.pivot,
				query: options.query,
				busy: options.busy,
				open: options.open,
				change: options.change,
				read: options.read,
			})}
		</div>`;
}

export function renderMarketplaceForm(
	session: AddInSession,
	busy: boolean,
	resolve: (event: SubmitEvent) => void,
) {
	if (!session.resolver || session.product === 'visio') return nothing;
	const { t } = session;
	return html`<form class="tools" @submit=${resolve}>
		<label for="marketplace-link">${t('Add from Marketplace link')}</label>
		<input id="marketplace-link" type="text" maxlength="2048" ?disabled=${busy} />
		<button type="submit" ?disabled=${busy}>${t('Review add-in')}</button>
	</form>`;
}
