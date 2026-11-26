export const OFFICE_ADD_IN_HOSTS = ['word', 'excel', 'powerpoint'] as const;
export type OfficeAddInHost = (typeof OFFICE_ADD_IN_HOSTS)[number];

export interface AddInRequirement {
	readonly kind: 'set' | 'method';
	readonly name: string;
	readonly minVersion: string;
	readonly maxVersion?: string;
}

export type AddInIssue =
	| { readonly kind: 'requirement'; readonly requirement: AddInRequirement }
	| { readonly kind: 'unsupported'; readonly feature: string };

export type AddInCommand = {
	readonly id: string;
	readonly label: string;
} & (
	| { readonly kind: 'taskPane'; readonly url: string }
	| { readonly kind: 'inactive'; readonly issue: AddInIssue }
);

export interface AddInTab {
	readonly id: string;
	readonly label: string;
	readonly groups: readonly {
		readonly id: string;
		readonly label: string;
		readonly commands: readonly AddInCommand[];
	}[];
}

export type AddInIcon =
	| { readonly kind: 'package'; readonly path: string }
	| { readonly kind: 'remote'; readonly url: string };

export interface AddInManifest {
	readonly format: 'xml' | 'unified';
	readonly id: string;
	readonly version: string;
	readonly publisher: string;
	readonly name: string;
	readonly description: string;
	readonly locale: string;
	readonly hosts: readonly OfficeAddInHost[];
	readonly requirements: readonly AddInRequirement[];
	readonly permissions: readonly string[];
	readonly entryOrigins: readonly string[];
	readonly declaredOrigins: readonly string[];
	readonly icons: readonly AddInIcon[];
	readonly identity: { readonly id: string; readonly resource?: string } | null;
	readonly pane: string | null;
	readonly tabs: readonly AddInTab[];
	readonly issues: readonly AddInIssue[];
}

export interface AddInReadOptions {
	readonly host: OfficeAddInHost;
	readonly locale?: string;
	/** The host's verified capabilities. The installation-only host advertises none. */
	readonly capabilities?: {
		readonly sets?: ReadonlyMap<string, string>;
		readonly methods?: ReadonlySet<string>;
	};
}

export interface LocalAddIn {
	readonly manifest: AddInManifest;
	readonly digest: string;
}
