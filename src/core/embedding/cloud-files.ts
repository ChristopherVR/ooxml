/**
 * Cloud file places for File > Open: a provider (OneDrive, or a host's own) that signs the user in,
 * lists their files, and opens one into an editor with a storage that saves back to that file.
 *
 * Format-neutral and DOM-free. Every product UI (Word, Excel, PowerPoint, Visio) and the Office
 * suite browse a provider through the shared `<office-ui-cloud-files>` element and load the
 * {@link CloudOpenedFile} it reports. Sign-in belongs to the host: a provider only calls back into
 * it, so how a token is obtained (an app registration, an existing session) is never product UI.
 */
import type { DocumentStorage } from './storage';

/** The Office formats a cloud place can open into an editor. */
export type CloudFileKind = 'docx' | 'xlsx' | 'pptx' | 'vsdx';

/** One row of a listing: a folder to browse into, or a file to open. */
export interface CloudEntry {
	/** Stable and unique within its provider (OneDrive: `driveId!itemId`). */
	readonly id: string;
	readonly name: string;
	readonly type: 'folder' | 'file';
	/** The Office format, for a file this library can open; absent for other files and folders. */
	readonly kind?: CloudFileKind;
	/** Size in bytes (files). */
	readonly size?: number;
	/** Last modified, epoch milliseconds. */
	readonly modified?: number;
	/** Display name of whoever modified it last. */
	readonly modifiedBy?: string;
	/** Human-readable place, for search and recent results (for example "My files > Reports"). */
	readonly location?: string;
	/** Number of children, for a folder, when the provider knows it. */
	readonly childCount?: number;
}

/** A page of entries. Pass `next` back as `page` for the following page. */
export interface CloudListing {
	readonly entries: readonly CloudEntry[];
	readonly next?: string;
}

/** The signed-in account, for display only. */
export interface CloudAccount {
	readonly name: string;
	readonly email?: string;
}

export interface CloudRequestOptions {
	readonly signal?: AbortSignal;
}

export interface CloudListOptions extends CloudRequestOptions {
	/** The `next` value of the previous page. */
	readonly page?: string;
}

export interface CloudOpenOptions extends CloudRequestOptions {
	/** Download progress, in bytes (`total` when the provider knows the size). */
	readonly onProgress?: (loaded: number, total?: number) => void;
}

/** A file opened from a cloud place, ready to load into the editor for its format. */
export interface CloudOpenedFile {
	readonly entry: CloudEntry;
	readonly kind: CloudFileKind;
	readonly fileName: string;
	readonly bytes: Uint8Array;
	/** False when the host's write policy (or the service) leaves the file read-only. */
	readonly canEdit: boolean;
	/**
	 * Saves to this same file (File > Save) or a sibling copy (Save As). The editor uses it for
	 * this document instead of its host `storage`; it resolves only once the provider stored the
	 * bytes and rejects on failure or an unresolvable conflict.
	 */
	readonly storage: DocumentStorage;
}

/**
 * A cloud place. Implemented by `createOneDriveFileProvider` (`ooxml-core/embedding/onedrive`) or
 * by a host for its own storage. Every method may reject; the UI shows the message.
 */
export interface CloudFileProvider {
	/** Stable id, for example `onedrive`. */
	readonly id: string;
	/** Display name, for example `OneDrive`. */
	readonly label: string;
	/** The signed-in account, or null when nobody is signed in. */
	account(): CloudAccount | null;
	/** Ask the host to sign the user in (it owns the sign-in UI). */
	signIn(): Promise<CloudAccount>;
	signOut?(): Promise<void>;
	/** Notifies when the account changes; returns the unsubscribe. */
	subscribe?(listener: () => void): () => void;
	/** The children of a folder; the root ("My files") when `folderId` is omitted. */
	list(folderId?: string, options?: CloudListOptions): Promise<CloudListing>;
	/** Recently used files, newest first. */
	recent?(options?: CloudListOptions): Promise<CloudListing>;
	/** Files matching `query` anywhere in the account. */
	search?(query: string, options?: CloudListOptions): Promise<CloudListing>;
	/** Download a file listed by this provider. */
	open(entry: CloudEntry, options?: CloudOpenOptions): Promise<CloudOpenedFile>;
	/** Open a file from a link the user pasted (a OneDrive sharing link, for example). */
	openLink?(link: string, options?: CloudOpenOptions): Promise<CloudOpenedFile>;
}

/** File extension to the Office format a cloud place opens, or undefined for anything else. */
export function cloudFileKind(name: string): CloudFileKind | undefined {
	const extension = /\.([a-z0-9]+)$/iu.exec(name)?.[1]?.toLowerCase();
	return extension === 'docx' ||
		extension === 'xlsx' ||
		extension === 'pptx' ||
		extension === 'vsdx'
		? extension
		: undefined;
}
