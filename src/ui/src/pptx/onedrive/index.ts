/** Compatibility entry: every product uses the same DOM-free OneDrive integration. */
export {
	createOneDriveClient,
	createOneDriveSession,
	encodeOneDriveSharingUrl,
	OneDriveError,
	OneDriveConflictError,
} from 'ooxml-core/embedding/onedrive';
export type {
	OneDriveClient,
	OneDriveClientOptions,
	OneDriveFile,
	OneDriveKind,
	OneDriveOpened,
	OneDriveRequestOptions,
	OneDriveShareOptions,
	OneDriveErrorCode,
	OneDriveConflict,
	OneDriveSession,
	OneDriveSessionSnapshot,
	OneDriveMergeInput,
	OneDriveMergeResult,
	OneDriveSessionOptions,
} from 'ooxml-core/embedding/onedrive';
