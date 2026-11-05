import { createStorageSession, createStorageMerger } from 'ooxml-core/embedding';
import {
	createOneDriveClient,
	createOneDriveFileProvider,
	createOneDriveSession,
	decodeOneDriveEntryId,
} from 'ooxml-core/embedding/onedrive';
import { officeMergeHandlers } from './onedrive-merge.js';
import { cloudIdentity } from './cloud-files.js';
import { resolveLaunchLinks } from './launch-link.js';

export const MICROSOFT_ACCOUNT = Object.freeze({ name: 'Microsoft account' });

/**
 * The browsable OneDrive place for the suite. Browsing, searching and sign-in come from
 * `createOneDriveFileProvider`; opening goes through the suite's own client so the opened file keeps
 * its identity and eTag (the suite saves through its cloud session, not the provider's storage).
 * Without a host `signIn` hook the host is treated as already signed in: its `getAccessToken` works.
 */
function oneDriveFileProvider(options, client, transfers) {
	const account = () =>
		typeof options.account === 'function' ? options.account() : options.account;
	const provider = createOneDriveFileProvider({
		...options,
		label: 'OneDrive',
		account: account() ?? (options.signIn ? null : MICROSOFT_ACCOUNT),
		signIn: options.signIn ?? (async () => account() ?? MICROSOFT_ACCOUNT),
		...(options.signOut ? { signOut: options.signOut } : {}),
	});
	const wrap = (entry, opened) => {
		const result = {
			entry,
			kind: opened.file.kind,
			fileName: opened.file.name,
			bytes: opened.bytes,
			canEdit: opened.file.access === 'write',
			// The suite saves through its own cloud session; this storage is never used.
			storage: {
				save: async () => {
					throw new Error('Save through the Office suite.');
				},
			},
		};
		transfers.set(result, opened);
		return result;
	};
	const transfer = (o) => ({
		...(o?.signal ? { signal: o.signal } : {}),
		...(o?.onProgress ? { onProgress: o.onProgress } : {}),
	});
	return {
		...provider,
		open: async (entry, o) => {
			if (entry.type !== 'file') throw new Error('Choose a file.');
			return wrap(entry, await client.openItem(decodeOneDriveEntryId(entry.id), transfer(o)));
		},
		openLink: async (link, o) => {
			const opened = await client.open(link, transfer(o));
			return wrap(
				{ id: link, name: opened.file.name, type: 'file', kind: opened.file.kind },
				opened,
			);
		},
	};
}

/** Resolve once before mounting. Provider credentials never enter persisted document metadata. */
export function resolveCloudOptions(host = globalThis) {
	const custom = host.ooxmlCloudOptions;
	if (
		custom &&
		(typeof custom.id !== 'string' ||
			!custom.id.trim() ||
			custom.id === 'onedrive' ||
			typeof custom.label !== 'string' ||
			!custom.label.trim() ||
			!custom.client ||
			typeof custom.serializeFile !== 'function' ||
			typeof custom.keyOf !== 'function')
	)
		throw new TypeError(
			'Custom cloud storage requires a unique id (other than onedrive), label, client, serializeFile and keyOf.',
		);
	const options = custom ?? host.ooxmlOneDriveOptions;
	const transfers = new WeakMap();
	const merge =
		options?.merge === false
			? undefined
			: (options?.merge ?? createStorageMerger({ ...officeMergeHandlers, ...options?.mergers }));
	const client =
		custom?.client ?? (options ? (options.client ?? createOneDriveClient(options)) : undefined);
	const fileProvider =
		custom?.fileProvider ??
		(!custom && options && !options.client && client
			? oneDriveFileProvider(options, client, transfers)
			: undefined);
	const id = custom?.id ?? 'onedrive';
	return {
		id,
		label: custom?.label ?? 'OneDrive',
		client,
		fileProvider,
		/** The `{ file, bytes }` a browser-opened file carries into the library. */
		sourceOf: (opened) => transfers.get(opened) ?? opened.cloudFile,
		createSession:
			options?.createSession ?? (custom ? createStorageSession : createOneDriveSession),
		serializeFile: custom?.serializeFile ?? cloudIdentity,
		keyOf: custom?.keyOf ?? ((file) => JSON.stringify([file.driveId, file.itemId])),
		merge,
		sessionOptions: {
			...options?.sessionOptions,
			...(options?.maxMergeAttempts === undefined
				? {}
				: { maxMergeAttempts: options.maxMergeAttempts }),
		},
		resolveConflict: options?.resolveConflict,
		shareChoices: options?.shareChoices,
		/** Opening `?open=<link>` launches (share target, web+ooxml:); null when off or unset. */
		launchLinks: options && client ? resolveLaunchLinks(options.launchLinks, id) : null,
	};
}
