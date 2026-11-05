import type {
	CloudAccount,
	CloudEntry,
	CloudFileKind,
	CloudFileProvider,
	CloudListOptions,
	CloudListing,
	CloudOpenOptions,
	CloudOpenedFile,
} from '../cloud-files';
import { decodeEntryId, encodeEntryId } from './browse';
import { createOneDriveClient } from './client';
import { createOneDriveSession } from './session';
import { OneDriveError } from './types';
import type {
	OneDriveClientOptions,
	OneDriveMergeInput,
	OneDriveMergeResult,
	OneDriveOpened,
} from './types';

export type OneDriveMerger = (
	input: OneDriveMergeInput,
) => OneDriveMergeResult | Promise<OneDriveMergeResult>;

export interface OneDriveFileProviderOptions extends OneDriveClientOptions {
	/** Display name; default `OneDrive`. */
	label?: string;
	/** The account already signed in when the provider is created. */
	account?: CloudAccount | null;
	/** The host's sign-in UI. Resolves with the account once `getAccessToken` can return a token. */
	signIn: () => Promise<CloudAccount>;
	signOut?: () => Promise<void>;
	/**
	 * Format-aware three-way mergers by kind (for example `mergeDocxSnapshots`), so concurrent
	 * edits to a file are combined on save. A kind without one reports a conflict instead.
	 */
	mergers?: Partial<Record<CloudFileKind, OneDriveMerger>>;
	/**
	 * Your own merge for every kind, which replaces `mergers`, or `false` to never merge: a save
	 * that meets a newer OneDrive copy then always reports a conflict for the host to resolve.
	 */
	merge?: OneDriveMerger | false;
	/** Maximum fresh-remote rebases per save when a merge applies. Default 3, maximum 10. */
	maxMergeAttempts?: number;
}

/** Lists, searches and opens a user's OneDrive files for File > Open (see `CloudFileProvider`). */
export function createOneDriveFileProvider(
	options: OneDriveFileProviderOptions,
): CloudFileProvider {
	const client = createOneDriveClient(options);
	const listeners = new Set<() => void>();
	let current: CloudAccount | null = options.account ?? null;
	const set = (account: CloudAccount | null) => {
		current = account;
		for (const listener of [...listeners]) listener();
	};
	const signedIn = () => {
		if (!current)
			throw new OneDriveError('authentication', 'Sign in to OneDrive to browse your files.');
	};
	const guard = <T>(run: () => Promise<T>): Promise<T> => {
		try {
			signedIn();
			return run();
		} catch (error) {
			return Promise.reject(error);
		}
	};
	const sessionFor = (opened: OneDriveOpened, entry: CloudEntry): CloudOpenedFile => {
		const merger =
			options.merge === false ? undefined : (options.merge ?? options.mergers?.[opened.file.kind]);
		const session = createOneDriveSession(client, opened, {
			...(merger ? { merge: merger } : {}),
			...(options.maxMergeAttempts === undefined
				? {}
				: { maxMergeAttempts: options.maxMergeAttempts }),
		});
		return {
			entry,
			kind: opened.file.kind,
			fileName: opened.file.name,
			bytes: opened.bytes,
			canEdit: opened.file.access === 'write',
			storage: session.storage,
		};
	};
	const request = (o?: CloudListOptions) => ({
		...(o?.signal ? { signal: o.signal } : {}),
		...(o?.page ? { page: o.page } : {}),
	});
	const transfer = (o?: CloudOpenOptions) => ({
		...(o?.signal ? { signal: o.signal } : {}),
		...(o?.onProgress ? { onProgress: o.onProgress } : {}),
	});
	return {
		id: 'onedrive',
		label: options.label ?? 'OneDrive',
		account: () => current,
		async signIn() {
			const account = await options.signIn();
			set(account);
			return account;
		},
		async signOut() {
			await options.signOut?.();
			set(null);
		},
		subscribe(listener) {
			listeners.add(listener);
			return () => void listeners.delete(listener);
		},
		list: (folderId, o): Promise<CloudListing> =>
			guard(() => client.list(folderId ? decodeEntryId(folderId) : undefined, request(o))),
		recent: (o) => guard(() => client.recent(request(o))),
		search: (query, o) => guard(() => client.search(query, request(o))),
		open: (entry, o) =>
			guard(async () => {
				if (entry.type !== 'file') throw new OneDriveError('unsupported', 'Choose a file.');
				const opened = await client.openItem(decodeEntryId(entry.id), transfer(o));
				return sessionFor(opened, entry);
			}),
		openLink: (link, o) =>
			guard(async () => {
				const opened = await client.open(link, transfer(o));
				const { file } = opened;
				return sessionFor(opened, {
					id: encodeEntryId(file),
					name: file.name,
					type: 'file',
					kind: file.kind,
					size: opened.bytes.length,
				});
			}),
	};
}
