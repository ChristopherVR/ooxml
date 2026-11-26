import { html, nothing, type TemplateResult } from 'lit';
import { glyph } from '../glyph';
import type { OfficeAddInOffer } from '../ribbon/add-in-settings';
import type { Pivot } from './add-ins-dialog';
import type { AddInSession } from './add-ins-session';
import { renderAddInDetails } from './add-ins-consent';
import { addInIssueLabels } from '../ribbon/add-in-local-dialog';

const MARKETPLACE = 'https://marketplace.microsoft.com/en-us/marketplace/apps?product=office';
export const LOCAL_INSTALL_REASON_ID = 'upload-unavailable';

function handoffUrl(link: string): string | undefined {
	try {
		const url = new URL(link);
		if (url.protocol === 'https:' && !url.username && !url.password) return url.href;
	} catch {
		// An invalid host configuration has no handoff.
	}
}

function tile(label: string, actions: unknown, details?: TemplateResult) {
	return html`<li class="tile"
		>${glyph('addIns', 'tile-icon')}<div class="tile-body">
			<p class="tile-name">${label}</p>${details}
			<div class="tile-actions">${actions}</div>
		</div></li
	>`;
}

interface ListOptions {
	session: AddInSession;
	pivot: Pivot;
	query: string;
	busy: boolean;
	open: (tab: () => void) => void;
	change: (action: () => void | Promise<void>) => void;
	read: (source: OfficeAddInOffer['read']) => void;
}

export function renderAddInList({ session, pivot, query, busy, open, change, read }: ListOptions) {
	const { installed, tabs, catalog, storefront, t, locale, product, launcher } = session;
	const filter = query.trim().toLocaleLowerCase(locale);
	const matches = (label: string) => label.toLocaleLowerCase(locale).includes(filter);
	const action = (label: string, name: string, run: () => void, disabledReason?: string) =>
		html`<button
			type="button"
			aria-label=${`${t(label)} ${name}`}
			?disabled=${busy || Boolean(disabledReason)}
			aria-describedby=${disabledReason ?? nothing}
			@click=${run}
			>${t(label)}</button
		>`;

	const installedIds = new Set(tabs.map(({ id }) => id));
	let tiles: TemplateResult[] = [];
	switch (pivot) {
		case 'My Add-ins':
			tiles = tabs
				.filter(({ label }) => matches(label))
				.map(({ id, label, open: tab }) => {
					const remove = catalog.find((entry) => entry.id === id)?.remove;
					return tile(label, [
						action('Open', label, () => open(tab)),
						remove ? action('Remove', label, () => change(remove)) : nothing,
					]);
				});
			for (const addIn of installed.values()) {
				const { manifest } = addIn;
				if (!matches(manifest.name)) continue;
				const { reason } = launcher.pane(addIn);
				const reasonId = `pane-unavailable-${manifest.id}`;
				tiles.push(
					tile(
						`${manifest.name} (${manifest.version})`,
						[
							action(
								'Open',
								manifest.name,
								() => open(() => void launcher.open(addIn, undefined, session.restoreFocus)),
								reason ? reasonId : undefined,
							),
							action('Remove', manifest.name, () => change(() => session.remove(manifest.id))),
						],
						html`${reason ? html`<p id=${reasonId}>${reason}</p>` : nothing}
							${addInIssueLabels(manifest, t).map((line) => html`<p>${line}</p>`)}
							<details
								><summary>${t('Add-in details')}</summary>
								${renderAddInDetails(addIn, t)}
							</details>`,
					),
				);
			}
			break;
		case 'Admin Managed':
			tiles = catalog
				.filter(({ id, label }) => !installedIds.has(id) && matches(label))
				.map(({ label, install }) =>
					tile(
						label,
						action('Install', label, () => change(install)),
					),
				);
			break;
		case 'Store':
			if (storefront === 'off') break;
			if (storefront.kind === 'catalog') {
				tiles = storefront.entries
					.filter(({ label }) => matches(label))
					.map((offer) =>
						tile(
							offer.label,
							action(
								'Install',
								offer.label,
								() => read((signal) => offer.read(signal)),
								product === 'visio' ? LOCAL_INSTALL_REASON_ID : undefined,
							),
						),
					);
				break;
			}
			const handoff = handoffUrl(storefront.url ?? MARKETPLACE);
			if (handoff) {
				return html`<p
						>${t('Browse add-ins for Microsoft Office. This editor runs add-ins supplied by its host.')}</p
					>
					<a href=${handoff} target="_blank" rel="noopener noreferrer"
						>${storefront.label ?? t('Browse Microsoft Marketplace')}</a
					>`;
			}
			break;
	}
	return tiles.length
		? html`<ul class="tiles"
				>${tiles}</ul
			>`
		: html`<p class="empty"
				>${t(filter ? 'No matching add-ins.' : 'No add-ins are available in this editor.')}</p
			>`;
}
