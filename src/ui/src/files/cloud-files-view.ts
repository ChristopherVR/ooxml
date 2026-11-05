import { html, nothing } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import type { CloudAccount, CloudEntry, CloudFileProvider } from 'ooxml-core/embedding';

import {
	entryBadge,
	fillLabel,
	formatModified,
	formatSize,
	type OfficeCloudFilesLabels,
} from './cloud-files-labels';

export type CloudFilesView = 'recent' | 'files' | 'search';

export interface CloudFilesCrumb {
	readonly id: string | undefined;
	readonly name: string;
}

/** What `<office-ui-cloud-files>` renders from: its state plus the actions the template calls. */
export interface CloudFilesViewModel {
	readonly text: OfficeCloudFilesLabels;
	readonly providers: readonly CloudFileProvider[];
	readonly locale: string | undefined;
	readonly view: CloudFilesView;
	readonly path: readonly CloudFilesCrumb[];
	readonly entries: readonly CloudEntry[];
	readonly next: string | undefined;
	readonly loading: boolean;
	readonly error: string;
	readonly query: string;
	readonly opening: string | undefined;
	readonly progress: { loaded: number; total?: number } | undefined;
	readonly focusIndex: number;
	accepted(entry: CloudEntry): boolean;
	select(index: number, entry: CloudEntry): void;
	show(view: CloudFilesView): void;
	enter(crumbIndex: number): void;
	loadMore(): void;
	choose(providerId: string): void;
	signIn(): void;
	signOut(): void;
	search(event: Event): void;
	openLink(event: Event): void;
	cancel(): void;
	keydown(event: KeyboardEvent): void;
}

function row(vm: CloudFilesViewModel, entry: CloudEntry, index: number) {
	const { text } = vm;
	const ok = vm.accepted(entry);
	const busy = vm.opening === entry.id;
	const blocked = entry.type === 'file' && (!ok || (vm.opening !== undefined && !busy));
	const place = vm.view === 'files' ? undefined : entry.location;
	return html`<li
		class="row"
		role="option"
		data-index=${index}
		data-entry=${entry.id}
		tabindex=${index === vm.focusIndex ? 0 : -1}
		aria-selected=${index === vm.focusIndex ? 'true' : 'false'}
		aria-disabled=${blocked ? 'true' : 'false'}
		aria-busy=${busy ? 'true' : 'false'}
		title=${ifDefined(entry.type === 'file' && !ok ? text.unsupported : undefined)}
		@click=${() => vm.select(index, entry)}
		><span class="name"
			><span
				class="icon ${entry.type === 'folder' ? 'folder' : (entry.kind ?? 'other')}"
				aria-hidden="true"
				>${entryBadge(entry)}</span
			><span class="title"
				><b>${entry.name}</b>${place ? html`<small>${place}</small>` : nothing}</span
			></span
		><span class="cell">${formatModified(entry.modified, vm.locale)}</span
		><span class="cell size">${entry.type === 'folder' ? '' : formatSize(entry.size)}</span></li
	>`;
}

function list(vm: CloudFilesViewModel) {
	const { text } = vm;
	if (vm.loading && vm.entries.length === 0)
		return html`<div aria-hidden="true"
			><div class="skeleton"></div><div class="skeleton"></div><div class="skeleton"></div
		></div>`;
	if (vm.entries.length === 0 && !vm.error)
		return html`<div class="empty">${vm.view === 'search' ? text.emptySearch : text.empty}</div>`;
	return html`
		<div class="head" aria-hidden="true"
			><span>${text.name}</span><span>${text.modified}</span
			><span class="size">${text.size}</span></div
		>
		<ul role="listbox" aria-label=${text.files} @keydown=${vm.keydown}
			>${vm.entries.map((entry, index) => row(vm, entry, index))}</ul
		>
		${
			vm.next
				? html`<button class="more" type="button" ?disabled=${vm.loading} @click=${vm.loadMore}
						>${text.loadMore}</button
					>`
				: nothing
		}
	`;
}

/** The signed-out state: the provider's name, an explanation and the host's Sign in. */
export function renderSignedOut(vm: CloudFilesViewModel, provider: CloudFileProvider) {
	const { text } = vm;
	const values = { provider: provider.label };
	return html`<div class="signin"
		><h2>${fillLabel(text.signInHeading, values)}</h2><p>${fillLabel(text.signInText, values)}</p
		><button type="button" class="primary" @click=${vm.signIn}>${text.signIn}</button
		><div class="error" role="alert">${vm.error}</div></div
	>`;
}

function toolbar(vm: CloudFilesViewModel, provider: CloudFileProvider, account: CloudAccount) {
	const { text } = vm;
	const tab = (view: CloudFilesView, label: string) =>
		html`<button
			type="button"
			data-view=${view}
			aria-pressed=${vm.view === view ? 'true' : 'false'}
			@click=${() => vm.show(view)}
			>${label}</button
		>`;
	const switcher =
		vm.providers.length > 1
			? html`<div class="tabs" role="group" aria-label=${text.providers}
					>${vm.providers.map(
						(other) =>
							html`<button
								type="button"
								data-provider=${other.id}
								aria-pressed=${other === provider ? 'true' : 'false'}
								@click=${() => vm.choose(other.id)}
								>${other.label}</button
							>`,
					)}</div
				>`
			: nothing;
	const search = provider.search
		? html`<office-ui-search
				placeholder=${text.searchPlaceholder}
				.value=${vm.query}
				@input=${vm.search}
			></office-ui-search>`
		: nothing;
	const signOut = provider.signOut
		? html`<button type="button" data-action="sign-out" @click=${vm.signOut}
				>${text.signOut}</button
			>`
		: nothing;
	return html`<div class="bar"
		>${switcher}<div class="tabs"
			>${provider.recent ? tab('recent', text.recent) : nothing}${tab('files', text.myFiles)}</div
		>${search}<span class="grow"></span><span class="account">${account.name}</span> ${signOut}</div
	>`;
}

function linkField(vm: CloudFilesViewModel) {
	const { text } = vm;
	return html`<form class="bar link" @submit=${vm.openLink}
		><input
			type="text"
			aria-label=${text.openLink}
			placeholder=${text.openLinkPlaceholder}
		/><button type="submit" ?disabled=${vm.opening !== undefined}
			>${text.openLinkButton}</button
		></form
	>`;
}

function statusLine(vm: CloudFilesViewModel) {
	const { text, progress } = vm;
	if (!vm.opening) return nothing;
	const entry = vm.entries.find((candidate) => candidate.id === vm.opening);
	const name = entry?.name ?? text.openLink;
	const percent =
		progress?.total && progress.total > 0
			? Math.min(100, Math.round((progress.loaded / progress.total) * 100))
			: undefined;
	const message =
		percent === undefined
			? fillLabel(text.opening, { name })
			: fillLabel(text.openingPercent, { name, percent: String(percent) });
	return html`<div class="status"
		><span role="status" data-progress=${ifDefined(percent)}>${message}</span
		><button type="button" data-action="cancel" @click=${vm.cancel}>${text.cancel}</button></div
	>`;
}

function crumbs(vm: CloudFilesViewModel) {
	return html`<nav aria-label=${vm.text.location}
		><ol
			>${vm.path.map(
				(crumb, index) =>
					html`<li
						><button
							type="button"
							data-crumb=${index}
							aria-current=${ifDefined(
								vm.view === 'files' && index === vm.path.length - 1 ? 'page' : undefined,
							)}
							@click=${() => vm.enter(index)}
							>${index === 0 ? vm.text.myFiles : crumb.name}</button
						></li
					>`,
			)}</ol
		></nav
	>`;
}

/** The signed-in browser: toolbar, optional link field, breadcrumb, listing and open status. */
export function renderBrowser(
	vm: CloudFilesViewModel,
	provider: CloudFileProvider,
	account: CloudAccount,
) {
	return html`<div aria-busy=${vm.loading ? 'true' : 'false'}
		>${toolbar(vm, provider, account)}${provider.openLink ? linkField(vm) : nothing}${
			vm.view === 'files' ? crumbs(vm) : nothing
		}${list(vm)}${statusLine(vm)} <div class="error" role="alert">${vm.error}</div></div
	>`;
}
