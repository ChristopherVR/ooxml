import { createStorageSession, createStorageMerger } from 'ooxml-core/embedding';
import { createOneDriveClient, createOneDriveSession } from 'ooxml-core/embedding/onedrive';
import { officeMergeHandlers } from './onedrive-merge.js';
import { cloudIdentity } from './cloud-files.js';

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
	const merge =
		options?.merge === false
			? undefined
			: (options?.merge ?? createStorageMerger({ ...officeMergeHandlers, ...options?.mergers }));
	return {
		id: custom?.id ?? 'onedrive',
		label: custom?.label ?? 'OneDrive',
		client:
			custom?.client ?? (options ? (options.client ?? createOneDriveClient(options)) : undefined),
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
	};
}
