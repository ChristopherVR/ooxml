import type { AddInDialogHost } from './add-in-dialog';

/** Trusted host commands for add-ins that are already implemented by the host. */
export interface OfficeAddInCatalogEntry {
	id: string;
	label: string;
	install: () => void | Promise<void>;
	remove?: () => void | Promise<void>;
}

/** An opt-in storefront offer. The dialog validates these bytes before asking for consent. */
export interface OfficeAddInOffer {
	id: string;
	label: string;
	read(signal: AbortSignal): File | Promise<File>;
}

export type OfficeAddInStorefront =
	| 'off'
	| { kind: 'microsoft'; url?: string; label?: string }
	| { kind: 'catalog'; entries: readonly OfficeAddInOffer[] };

/** The callback receives only a normalized Marketplace asset ID. */
export type OfficeAddInMarketplaceResolver = (
	assetId: string,
	signal: AbortSignal,
) => Promise<Uint8Array | ArrayBuffer>;

const CATALOG = Symbol.for('office-ui.add-in-catalog');
const STOREFRONT = Symbol.for('office-ui.add-in-storefront');
const RESOLVER = Symbol.for('office-ui.add-in-marketplace-resolver');
const ENTRY_LIMIT = 128;

type HostSettings = AddInDialogHost & {
	[CATALOG]?: readonly OfficeAddInCatalogEntry[];
	[STOREFRONT]?: OfficeAddInStorefront;
	[RESOLVER]?: OfficeAddInMarketplaceResolver | null;
};

/** Set or clear the trusted catalog for a single editor or host container. */
export function setOfficeAddInCatalog(
	host: AddInDialogHost,
	entries: readonly OfficeAddInCatalogEntry[],
): void {
	const target: HostSettings = host;
	if (entries.length) target[CATALOG] = entries.slice(0, ENTRY_LIMIT);
	else delete target[CATALOG];
}

/** Enable or disable a storefront for one editor or host container. The default is off. */
export function setOfficeAddInStorefront(
	host: AddInDialogHost,
	storefront: OfficeAddInStorefront,
): void {
	const target: HostSettings = host;
	target[STOREFRONT] =
		storefront !== 'off' && storefront.kind === 'catalog'
			? { kind: 'catalog', entries: storefront.entries.slice(0, ENTRY_LIMIT) }
			: storefront;
}

/** Register a resolver, or pass null to turn resolution off inside this scope. */
export function setOfficeAddInMarketplaceResolver(
	host: AddInDialogHost,
	resolver: OfficeAddInMarketplaceResolver | null,
): void {
	const target: HostSettings = host;
	target[RESOLVER] = resolver;
}

export function addInSettings(scopes: readonly AddInDialogHost[]) {
	let catalog: readonly OfficeAddInCatalogEntry[] | undefined;
	let storefront: OfficeAddInStorefront | undefined;
	let resolver: OfficeAddInMarketplaceResolver | null | undefined;
	for (const current of scopes) {
		const settings: HostSettings = current;
		catalog ??= settings[CATALOG];
		storefront ??= settings[STOREFRONT];
		if (resolver === undefined) resolver = settings[RESOLVER];
	}
	return { catalog: catalog ?? [], storefront: storefront ?? 'off', resolver };
}
