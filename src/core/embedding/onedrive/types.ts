import type { DocumentStorage } from '../storage';

export type OneDriveKind = 'docx' | 'xlsx' | 'pptx' | 'vsdx';
/** Safe persisted identity. Never contains access tokens or temporary transfer URLs. */
export interface OneDriveFile {
	driveId: string;
	itemId: string;
	parentId?: string;
	name: string;
	kind: OneDriveKind;
	eTag: string;
	webUrl?: string;
	access: 'read' | 'write' | 'unknown';
}
export interface OneDriveOpened {
	file: OneDriveFile;
	bytes: Uint8Array;
}
export interface OneDriveRequestOptions {
	signal?: AbortSignal;
}
export interface OneDriveShareOptions {
	type: 'view' | 'edit';
	scope: 'anonymous' | 'organization' | 'users';
}
export interface OneDriveClient {
	open(shareUrl: string, options?: OneDriveRequestOptions): Promise<OneDriveOpened>;
	read(file: OneDriveFile, options?: OneDriveRequestOptions): Promise<OneDriveOpened>;
	save(
		file: OneDriveFile,
		bytes: Uint8Array,
		options?: OneDriveRequestOptions,
	): Promise<OneDriveFile>;
	saveCopy(
		file: OneDriveFile,
		bytes: Uint8Array,
		name?: string,
		options?: OneDriveRequestOptions,
	): Promise<OneDriveOpened>;
	share(
		file: OneDriveFile,
		link: OneDriveShareOptions,
		options?: OneDriveRequestOptions,
	): Promise<string>;
}
export interface OneDriveClientOptions {
	getAccessToken: () => Promise<string>;
	fetch?: typeof globalThis.fetch;
	maxBytes?: number;
	timeoutMs?: number;
	/** Trusted host policy; token scopes or a sharing link alone do not establish edit access. */
	canWrite?: (file: Readonly<OneDriveFile>) => boolean | Promise<boolean>;
}
export type OneDriveErrorCode =
	| 'invalid-link'
	| 'authentication'
	| 'access-denied'
	| 'conflict'
	| 'locked'
	| 'unsupported'
	| 'too-large'
	| 'invalid-response'
	| 'network';
export class OneDriveError extends Error {
	constructor(
		public readonly code: OneDriveErrorCode,
		message: string,
		public readonly status?: number,
	) {
		super(message);
		this.name = 'OneDriveError';
	}
}
export interface OneDriveConflict {
	base: OneDriveOpened;
	local: Uint8Array;
	remote: OneDriveOpened;
	/** Unresolved merge diagnostics. These never contain a caught provider exception. */
	issues?: readonly string[];
}
export class OneDriveConflictError extends OneDriveError {
	constructor(public readonly conflict: OneDriveConflict) {
		super(
			'conflict',
			'The changes could not be combined automatically. Your edits are retained for resolution.',
		);
		this.name = 'OneDriveConflictError';
	}
}
export interface OneDriveMergeInput extends OneDriveConflict {
	kind: OneDriveKind;
	signal?: AbortSignal;
}
export type OneDriveMergeResult =
	| { status: 'merged'; bytes: Uint8Array }
	| { status: 'unresolved'; issues: readonly string[] };
export interface OneDriveSessionOptions {
	/** A format-aware three-way merger. No generic ZIP or last-writer fallback is provided. */
	merge?: (input: OneDriveMergeInput) => OneDriveMergeResult | Promise<OneDriveMergeResult>;
	/** Maximum fresh-remote rebases per save when merge is supplied. Default 3, maximum 10. */
	maxMergeAttempts?: number;
}
export interface OneDriveSessionSnapshot {
	file: OneDriveFile;
	baseBytes: Uint8Array;
	conflict: OneDriveConflict | null;
}
export interface OneDriveSession {
	getSnapshot(): OneDriveSessionSnapshot;
	save(bytes: Uint8Array, options?: OneDriveRequestOptions): Promise<OneDriveFile>;
	reload(options?: OneDriveRequestOptions): Promise<OneDriveOpened>;
	saveCopy(
		bytes?: Uint8Array,
		name?: string,
		options?: OneDriveRequestOptions,
	): Promise<OneDriveOpened>;
	storage: DocumentStorage;
}
