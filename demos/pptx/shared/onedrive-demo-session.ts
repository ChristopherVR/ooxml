/**
 * The OneDrive demo's file session: open a sharing link, hand the viewer one stable `storage`
 * object, and save, merge or recover through `ooxml-core/embedding/onedrive`. Framework-free; each
 * demo maps the deck it reports onto its own viewer.
 */
import type { DocumentStorage } from 'ooxml-core/embedding';
import {
	createOneDriveClient,
	createOneDriveSession,
	OneDriveConflictError,
	type OneDriveClient,
	type OneDriveOpened,
	type OneDriveSession,
} from 'ooxml-core/embedding/onedrive';
import { mergePptxSnapshots } from 'ooxml-core/pptx';

import type { OneDriveSignIn } from '../../onedrive/auth';

/** What the viewer shows: load `bytes` whenever `revision` changes. */
export interface OneDriveDemoDeck {
	readonly bytes: Uint8Array;
	readonly fileName: string;
	readonly canEdit: boolean;
	/** Bumped for every new file and every replacement (a merge, or "open latest"). */
	readonly revision: number;
}

export interface OneDriveDemoEvents {
	deck(deck: OneDriveDemoDeck): void;
	status(text: string, tone?: 'ok' | 'busy' | 'error'): void;
	/** Overlapping edits: the panel offers "save a copy" and "open latest". */
	conflict(open: boolean): void;
}

const sameBytes = (a: Uint8Array, b: Uint8Array) =>
	a.length === b.length && a.every((value, index) => value === b[index]);

const message = (error: unknown) => (error instanceof Error ? error.message : String(error));

export interface OneDriveDemoSession {
	/** Pass this to the viewer once; it always saves to the file opened last. */
	readonly storage: DocumentStorage;
	open(link: string): Promise<void>;
	saveMyCopy(): Promise<void>;
	openLatest(): Promise<void>;
	/** Sign-in changed: the next open reads the newly signed-in drive. */
	forgetSignIn(): void;
}

export function createOneDriveDemoSession(
	signIn: () => OneDriveSignIn | null,
	events: OneDriveDemoEvents,
): OneDriveDemoSession {
	let drive: OneDriveClient | null = null;
	let ownDriveId: string | null = null;
	let session: OneDriveSession | null = null;
	let conflict: OneDriveConflictError | null = null;
	let revision = 0;

	const show = (opened: OneDriveOpened, canEdit: boolean) => {
		revision += 1;
		events.deck({ bytes: opened.bytes, fileName: opened.file.name, canEdit, revision });
	};
	const setConflict = (next: OneDriveConflictError | null) => {
		conflict = next;
		events.conflict(next !== null);
	};

	async function client(): Promise<OneDriveClient> {
		const account = signIn();
		if (!account) {
			throw new Error('Sign in first.');
		}
		if (!drive) {
			const getAccessToken = () => account.getAccessToken();
			// Edit policy for the demo: only files on the signed-in user's own drive are editable.
			const response = await fetch('https://graph.microsoft.com/v1.0/me/drive?$select=id', {
				headers: { Authorization: `Bearer ${await getAccessToken()}` },
				credentials: 'omit',
			});
			if (!response.ok) {
				throw new Error(`Could not read your OneDrive (HTTP ${response.status}). Sign in again.`);
			}
			ownDriveId = ((await response.json()) as { id?: string }).id ?? null;
			drive = createOneDriveClient({
				getAccessToken,
				canWrite: (file) => ownDriveId !== null && file.driveId === ownDriveId,
			});
		}
		return drive;
	}

	const storage: DocumentStorage = {
		async save(bytes, meta) {
			if (!session) {
				throw new Error('Open a OneDrive file first.');
			}
			events.status('Saving to OneDrive...', 'busy');
			try {
				const result = await session.storage.save(bytes, meta);
				const saved = session.getSnapshot();
				// Someone else's edits were merged in: show the combined deck.
				if (!sameBytes(saved.baseBytes, bytes)) {
					show({ file: saved.file, bytes: saved.baseBytes }, saved.file.access === 'write');
					events.status('Saved to OneDrive, with the other edits merged in', 'ok');
				} else {
					events.status(`Saved to OneDrive at ${new Date().toLocaleTimeString()}`, 'ok');
				}
				return result;
			} catch (error) {
				if (error instanceof OneDriveConflictError) {
					setConflict(error);
					events.status('Someone else changed the same content. Choose how to continue.', 'error');
				} else {
					events.status(`Not saved: ${message(error)}`, 'error');
				}
				throw error;
			}
		},
	};

	return {
		storage,
		async open(link) {
			events.status('Opening...', 'busy');
			try {
				const opened = await (await client()).open(link);
				if (opened.file.kind !== 'pptx') {
					throw new Error(`This is a .${opened.file.kind} file; the PowerPoint demo opens .pptx.`);
				}
				session = createOneDriveSession(drive!, opened, {
					merge: ({ base, local, remote }) => mergePptxSnapshots(base.bytes, local, remote.bytes),
				});
				setConflict(null);
				const canEdit = opened.file.access === 'write';
				show(opened, canEdit);
				events.status(
					canEdit
						? 'Opened from OneDrive. Save (Ctrl+S) uploads to the same file.'
						: 'Opened read-only: the demo only edits files on your own OneDrive.',
					'ok',
				);
			} catch (error) {
				events.status(`Could not open: ${message(error)}`, 'error');
			}
		},
		async saveMyCopy() {
			if (!session || !conflict) {
				return;
			}
			const name = session.getSnapshot().file.name.replace(/\.pptx$/iu, ' (my edits).pptx');
			const copy = await session.saveCopy(conflict.conflict.local, name);
			setConflict(null);
			show(copy, copy.file.access === 'write');
			events.status(`Saved your edits as ${copy.file.name}; the original is unchanged.`, 'ok');
		},
		async openLatest() {
			if (!session) {
				return;
			}
			const latest = await session.reload();
			setConflict(null);
			show(latest, latest.file.access === 'write');
			events.status('Opened the latest OneDrive version; your unsaved edits were discarded.', 'ok');
		},
		forgetSignIn() {
			drive = null;
			ownDriveId = null;
		},
	};
}
