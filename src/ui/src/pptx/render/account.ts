import type { DocumentStorage } from 'ooxml-core/embedding';

import { clearAutosaveSnapshots } from './autosave-maintenance';
import { listAutosaveSnapshots } from './autosave-store';
import { clearStoredViewerPrefs, writeStoredViewerPrefs } from './viewer-prefs-storage';

/** Local-only display profile shown in File > Account. Never sent anywhere; purely cosmetic. */
export interface ViewerProfile {
	displayName: string;
	/** CSS color for the avatar bubble background, e.g. `'#c2431f'`. */
	avatarColor: string;
	/** Single character shown in the avatar bubble. Derived from `displayName` when omitted. */
	initial?: string;
}

export const DEFAULT_VIEWER_PROFILE: ViewerProfile = {
	displayName: '',
	avatarColor: '#6366f1',
};

/** Suggested avatar color swatches for the Account profile editor. */
export const AVATAR_COLOR_SWATCHES: readonly string[] = [
	'#6366f1', // indigo
	'#c2431f', // vermilion
	'#0891b2', // cyan
	'#16a34a', // green
	'#ca8a04', // amber
	'#db2777', // pink
	'#64748b', // slate
];

/** Derive the avatar-bubble initial from a profile, falling back to `'?'` for an empty name. */
export function resolveProfileInitial(profile: ViewerProfile): string {
	if (profile.initial) {
		return profile.initial.slice(0, 1).toUpperCase();
	}
	const trimmed = profile.displayName.trim();
	return trimmed ? trimmed.slice(0, 1).toUpperCase() : '?';
}

/**
 * Optional hook point for hosts that want to wire a real sign-in flow into
 * File > Account. Disabled by default: the Account page renders nothing
 * extra unless a host explicitly opts in by passing `enabled: true`.
 * See docs/guide for wiring instructions.
 */
export interface AccountAuthConfig {
	enabled: boolean;
	onSignIn: () => void;
	signedInUser?: {
		name: string;
		email?: string;
		avatarUrl?: string;
	};
}

export interface LocalStorageUsageSummary {
	/** Number of presentations with a local autosave/recovery snapshot. */
	presentationCount: number;
	/** Total bytes across every stored snapshot. */
	totalBytes: number;
}

/**
 * Summarize recovery-snapshot usage for the Account > Storage & Privacy panel: the library's
 * IndexedDB store, or the host's recovery store when the viewer's `storage` has one (empty when
 * that host does not list its copies).
 */
export async function getLocalStorageUsageSummary(
	storage?: DocumentStorage | null,
): Promise<LocalStorageUsageSummary> {
	const snapshots = await listAutosaveSnapshots(storage);
	return {
		presentationCount: snapshots.length,
		totalBytes: snapshots.reduce((sum, snapshot) => sum + snapshot.size, 0),
	};
}

/**
 * Delete every locally stored presentation/recovery snapshot and persisted
 * viewer preference (theme/locale/profile). Irreversible; callers should
 * confirm with the user before invoking this.
 */
export async function clearAllLocalViewerData(storage?: DocumentStorage | null): Promise<void> {
	// With a host recovery store, only one that lists its copies is cleared (see autosave-maintenance).
	await clearAutosaveSnapshots(storage);
	clearStoredViewerPrefs();
}

/** Persist an updated profile via the shared viewer-prefs store. */
export function saveViewerProfile(profile: ViewerProfile): void {
	writeStoredViewerPrefs({ profile });
}
