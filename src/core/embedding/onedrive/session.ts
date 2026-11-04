import { createStorageSession } from '../storage-session';
import { officeKind } from './metadata';
import { OneDriveConflictError, OneDriveError } from './types';
import type {
	OneDriveClient,
	OneDriveOpened,
	OneDriveSession,
	OneDriveSessionOptions,
} from './types';

/** OneDrive transport adapter over the provider-neutral serialized save/merge session. */
export function createOneDriveSession(
	client: OneDriveClient,
	opened: OneDriveOpened,
	settings: OneDriveSessionOptions = {},
): OneDriveSession {
	return createStorageSession(
		{
			read: (file, options) => client.read(file, options),
			save: (file, bytes, options) => client.save(file, bytes, options),
			saveCopy: (file, bytes, name, options) => client.saveCopy(file, bytes, name, options),
			isConflict: (error) => error instanceof OneDriveError && error.code === 'conflict',
		},
		opened,
		{
			...settings,
			conflictError: (value) => new OneDriveConflictError(value),
			validateSave(meta, file) {
				if (officeKind(meta.fileName) !== file.kind)
					throw new OneDriveError(
						'unsupported',
						'A different Office format needs a separate export destination.',
					);
			},
		},
	);
}
