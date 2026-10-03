import { tok } from './tokens.js';
import { initialsOf } from './presence.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

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

const CSS = `
:host { display: block; font-family: ${tok('--office-font')}; font-size: ${tok('--office-font-size')}; }
h2 { margin: 0 0 ${tok('--office-space-3')}; font-size: ${tok('--office-font-size-lg')}; font-weight: 600; }
.who { display: flex; align-items: center; gap: ${tok('--office-space-3')}; margin-bottom: ${tok('--office-space-3')}; }
.avatar { display: grid; place-items: center; width: ${tok('--office-avatar-size')}; height: ${tok('--office-avatar-size')}; border-radius: 50%;
	color: ${tok('--office-accent-foreground')}; font-weight: ${tok('--office-font-weight-bold')}; font-size: ${tok('--office-font-size-xl')}; flex: none; }
.name { font-size: ${tok('--office-font-size-md')}; font-weight: 600; }
.note { margin: ${tok('--office-space-0')} 0 0; color: ${tok('--office-muted-foreground')}; font-size: ${tok('--office-font-size-sm')}; }
label { display: grid; gap: ${tok('--office-space-1')}; max-width: ${tok('--office-account-field-width')}; margin-bottom: ${tok('--office-space-3')}; }
input { min-height: ${tok('--office-field-height')}; box-sizing: border-box; padding: ${tok('--office-space-0')} ${tok('--office-space-1-5')}; font: inherit; color: inherit;
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius-sm')}; background: ${tok('--office-background')}; }
.swatches { display: flex; flex-wrap: wrap; gap: ${tok('--office-space-1-5')}; margin: ${tok('--office-space-1')} 0 ${tok('--office-space-3')}; }
.swatches button { width: ${tok('--office-target-size')}; height: ${tok('--office-target-size')}; padding: 0;
	border: ${tok('--office-border-width-thick')} solid transparent; border-radius: 50%; cursor: pointer; }
.swatches button[aria-checked="true"] { border-color: ${tok('--office-foreground')};
	box-shadow: inset 0 0 0 ${tok('--office-border-width-thick')} ${tok('--office-background')}; }
.swatches button:focus-visible, input:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: calc(${tok('--office-focus-offset')} / 2); }
@media (forced-colors: active) { .swatches button { forced-color-adjust: none; } }
`;

/**
 * `<office-ui-account>`: User Information with avatar, display name and avatar colour. Set or read
 * `profile`; edits emit `office-profile-change`. `slot="sign-in"` hosts a product sign-in flow;
 * the default slot holds product sections (theme, product information).
 */
export const defineAccount = definer('office-ui-account', () => {
	class OfficeUiAccount extends HTMLElement {
		#profile: OfficeProfile = { ...DEFAULT_OFFICE_PROFILE };
		readonly #avatar: HTMLElement;
		readonly #name: HTMLElement;
		readonly #input: HTMLInputElement;
		readonly #swatches: HTMLElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			const heading = doc.createElement('h2');
			heading.textContent = 'User Information';
			const who = doc.createElement('div');
			who.className = 'who';
			this.#avatar = doc.createElement('div');
			this.#avatar.className = 'avatar';
			this.#avatar.setAttribute('aria-hidden', 'true');
			const details = doc.createElement('div');
			this.#name = doc.createElement('div');
			this.#name.className = 'name';
			const note = doc.createElement('p');
			note.className = 'note';
			note.textContent = 'Stored on this device only. Never sent anywhere.';
			details.append(this.#name, note);
			who.append(this.#avatar, details);
			const field = doc.createElement('label');
			field.append('Display name');
			this.#input = doc.createElement('input');
			this.#input.type = 'text';
			this.#input.maxLength = 64;
			this.#input.autocomplete = 'name';
			this.#input.addEventListener('change', () =>
				this.#update({ ...this.#profile, displayName: this.#input.value.trim() }),
			);
			field.append(this.#input);
			const colorLabel = doc.createElement('div');
			colorLabel.id = 'color-label';
			colorLabel.textContent = 'Avatar colour';
			this.#swatches = doc.createElement('div');
			this.#swatches.className = 'swatches';
			this.#swatches.setAttribute('role', 'radiogroup');
			this.#swatches.setAttribute('aria-labelledby', 'color-label');
			for (const color of OFFICE_AVATAR_SWATCHES) {
				const swatch = doc.createElement('button');
				swatch.type = 'button';
				swatch.setAttribute('role', 'radio');
				swatch.setAttribute('aria-label', color);
				swatch.dataset.color = color;
				swatch.style.background = color;
				swatch.addEventListener('click', () =>
					this.#update({ ...this.#profile, avatarColor: color }),
				);
				this.#swatches.append(swatch);
			}
			this.#swatches.addEventListener('keydown', (event) => this.#swatchKey(event));
			const signIn = doc.createElement('slot');
			signIn.name = 'sign-in';
			root.append(
				heading,
				who,
				field,
				colorLabel,
				this.#swatches,
				signIn,
				doc.createElement('slot'),
			);
			this.#sync();
		}
		get profile(): OfficeProfile {
			return { ...this.#profile };
		}
		set profile(value: OfficeProfile) {
			this.#profile = sanitizeOfficeProfile(value);
			this.#sync();
		}
		#update(next: OfficeProfile): void {
			this.#profile = sanitizeOfficeProfile(next);
			this.#sync();
			emit(this, 'office-profile-change', { profile: this.profile });
		}
		#swatchKey(event: KeyboardEvent): void {
			const step = { ArrowRight: 1, ArrowDown: 1, ArrowLeft: -1, ArrowUp: -1 }[event.key];
			if (!step) return;
			event.preventDefault();
			const colors = OFFICE_AVATAR_SWATCHES;
			const at = Math.max(0, colors.indexOf(this.#profile.avatarColor));
			const color = colors[(at + step + colors.length) % colors.length]!;
			this.#update({ ...this.#profile, avatarColor: color });
			this.#swatches.querySelector<HTMLButtonElement>(`[data-color="${color}"]`)?.focus();
		}
		#sync(): void {
			const { displayName, avatarColor } = this.#profile;
			this.#avatar.textContent = profileInitials(this.#profile);
			this.#avatar.style.background = avatarColor;
			this.#name.textContent = displayName || 'No display name';
			if (this.#input.value.trim() !== displayName) this.#input.value = displayName;
			const known = OFFICE_AVATAR_SWATCHES.includes(avatarColor);
			this.#swatches
				.querySelectorAll<HTMLButtonElement>('[role="radio"]')
				.forEach((swatch, index) => {
					const checked = swatch.dataset.color === avatarColor;
					swatch.setAttribute('aria-checked', String(checked));
					swatch.tabIndex = checked || (!known && index === 0) ? 0 : -1;
				});
		}
	}
	return OfficeUiAccount;
});
