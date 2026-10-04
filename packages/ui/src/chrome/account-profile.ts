import { initialsOf } from '../presence.js';

/**
 * Office File > Account: the local display profile (name and avatar colour) shared by the suite.
 * Moved from pptx-viewer `render/account.ts`. The profile is cosmetic and local: it is stored on
 * this device only and never sent anywhere, except as the collaborator identity a product passes
 * to a live session the user starts.
 */
export interface OfficeProfile {
	displayName: string;
	/** CSS colour of the avatar bubble, e.g. `'#c2431f'`. */
	avatarColor: string;
	/** Character(s) in the avatar bubble; derived from `displayName` when omitted. */
	initial?: string;
}

export const DEFAULT_OFFICE_PROFILE: OfficeProfile = { displayName: '', avatarColor: '#6366f1' };

/** Suggested avatar colours (indigo, vermilion, cyan, green, amber, pink, slate). */
export const OFFICE_AVATAR_SWATCHES: readonly string[] = [
	'#6366f1',
	'#c2431f',
	'#0891b2',
	'#16a34a',
	'#ca8a04',
	'#db2777',
	'#64748b',
];

/** One key for the whole suite, so Word, PowerPoint, Excel and Visio on one origin share it. */
export const OFFICE_PROFILE_STORAGE_KEY = 'ooxml-office-profile';

export type OfficeProfileChangeEvent = CustomEvent<{ profile: OfficeProfile }>;

const COLOR = /^#[0-9a-f]{6}$/i;

/** Validate untrusted stored data into a profile. */
export function sanitizeOfficeProfile(raw: unknown): OfficeProfile {
	const value = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>;
	const displayName =
		typeof value.displayName === 'string' ? value.displayName.trim().slice(0, 64) : '';
	const avatarColor =
		typeof value.avatarColor === 'string' && COLOR.test(value.avatarColor)
			? value.avatarColor
			: DEFAULT_OFFICE_PROFILE.avatarColor;
	const initial =
		typeof value.initial === 'string' && value.initial.trim()
			? value.initial.trim().slice(0, 2)
			: undefined;
	return { displayName, avatarColor, ...(initial ? { initial } : {}) };
}

/** The avatar bubble text: explicit initial, else initials of the name, else `?`. */
export function profileInitials(profile: OfficeProfile): string {
	return (profile.initial ?? initialsOf(profile.displayName)).toUpperCase();
}

type StorageLike = Pick<Storage, 'getItem' | 'setItem' | 'removeItem'>;
const defaultStorage = (): StorageLike | undefined => {
	try {
		return typeof localStorage === 'undefined' ? undefined : localStorage;
	} catch {
		return undefined;
	}
};

/** Read the stored profile; the default when storage is unavailable, empty or corrupt. */
export function readOfficeProfile(
	storage: StorageLike | undefined = defaultStorage(),
	key = OFFICE_PROFILE_STORAGE_KEY,
): OfficeProfile {
	try {
		const raw = storage?.getItem(key);
		return raw ? sanitizeOfficeProfile(JSON.parse(raw)) : { ...DEFAULT_OFFICE_PROFILE };
	} catch {
		return { ...DEFAULT_OFFICE_PROFILE };
	}
}

/** Persist the profile on this device. Silently does nothing when storage is unavailable. */
export function writeOfficeProfile(
	profile: OfficeProfile,
	storage: StorageLike | undefined = defaultStorage(),
	key = OFFICE_PROFILE_STORAGE_KEY,
): void {
	try {
		storage?.setItem(key, JSON.stringify(sanitizeOfficeProfile(profile)));
	} catch {
		/* Private browsing or quota: the profile stays for this session only. */
	}
}

/** Forget the stored profile. */
export function clearOfficeProfile(
	storage: StorageLike | undefined = defaultStorage(),
	key = OFFICE_PROFILE_STORAGE_KEY,
): void {
	try {
		storage?.removeItem(key);
	} catch {
		/* Nothing stored. */
	}
}
