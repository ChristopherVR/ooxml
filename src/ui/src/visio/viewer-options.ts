import {
	readOfficeProfile,
	sanitizeOfficeProfile,
	writeOfficeProfile,
	type OfficeOptionCategory,
	type OfficeOptionValues,
	type OfficeOptionsChangeEvent,
	type OfficeProfile,
	type OfficeProfileChangeEvent,
} from '../controls';
import { VISIO_OPTION_CATEGORIES } from './viewer-options-categories';

export { VISIO_OPTION_CATEGORIES };

type OptionsDialog = HTMLElement & {
	categories: readonly OfficeOptionCategory[];
	values: OfficeOptionValues;
	defaults: OfficeOptionValues | null;
	category: string;
	show(): void;
};
type Account = HTMLElement & { profile: OfficeProfile };

/** The status bar's look: Visio's product colour ("Colorful") or the neutral suite bar. */
export type StatusBarTheme = 'colorful' | 'neutral';
const THEME_KEY = 'ooxml-visio-office-theme';
export const isStatusBarTheme = (value: unknown): value is StatusBarTheme =>
	value === 'colorful' || value === 'neutral';
/** The Office Theme remembered on this device, if any (storage may be unavailable). */
export function readStatusBarTheme(): StatusBarTheme | undefined {
	try {
		const stored = globalThis.localStorage?.getItem(THEME_KEY);
		return isStatusBarTheme(stored) ? stored : undefined;
	} catch {
		return undefined;
	}
}
function writeStatusBarTheme(theme: StatusBarTheme): void {
	try {
		globalThis.localStorage?.setItem(THEME_KEY, theme);
	} catch {
		// Kept for this session only.
	}
}

/** Office Theme as Options offers it: the two light looks, or the dark theme. */
export type OfficeThemeChoice = StatusBarTheme | 'dark';
const isOfficeThemeChoice = (value: unknown): value is OfficeThemeChoice =>
	value === 'dark' || isStatusBarTheme(value);

function profileValues(profile: OfficeProfile, officeTheme: OfficeThemeChoice): OfficeOptionValues {
	return {
		officeTheme,
		miniToolbar: false,
		livePreview: false,
		screenTips: 'descriptions',
		undoLevels: 20,
		autoSizeText: false,
		userName: profile.displayName,
		userInitials: profile.initial ?? '',
	};
}

/** The Options dialog element. Built once and appended to the viewer's shadow root. */
export function createOptionsDialog(doc: Document): HTMLElement {
	const dialog = doc.createElement('office-ui-options-dialog') as OptionsDialog;
	dialog.setAttribute('heading', 'Visio Options');
	dialog.categories = VISIO_OPTION_CATEGORIES;
	// Reset all options: no user name or initials, and the neutral status bar.
	dialog.defaults = { userName: '', userInitials: '', officeTheme: 'neutral' };
	return dialog;
}

/**
 * File > Account and File > Options share one local profile. Edits in either are written to the
 * suite-wide profile store, mirrored into the other and reported through `onProfile`.
 */
export class ViewerProfile {
	#profile: OfficeProfile = readOfficeProfile();
	readonly #dialog: OptionsDialog;
	readonly #account: Account | null;
	constructor(
		root: ShadowRoot,
		private readonly onProfile: (profile: OfficeProfile) => void = () => {},
		/** Reads and applies the Office Theme on the viewer element. */
		private readonly theme: { get(): OfficeThemeChoice; set(theme: OfficeThemeChoice): void } = {
			get: () => 'neutral',
			set: () => {},
		},
	) {
		this.#dialog = root.querySelector<OptionsDialog>('office-ui-options-dialog')!;
		this.#account = root.querySelector<Account>('office-ui-account');
		this.#sync();
	}
	get profile(): OfficeProfile {
		return { ...this.#profile };
	}
	/** Open Visio Options, on General as Visio does. */
	showOptions(category = 'general'): void {
		this.#dialog.values = profileValues(this.#profile, this.theme.get());
		this.#dialog.category = category;
		this.#dialog.show();
	}
	wire(): () => void {
		const Abort = this.#dialog.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.#dialog.addEventListener(
			'office-options-change',
			(event) => {
				const { values } = (event as OfficeOptionsChangeEvent).detail;
				if (isOfficeThemeChoice(values.officeTheme) && values.officeTheme !== this.theme.get()) {
					this.theme.set(values.officeTheme);
					// The light looks are remembered here; light or dark is the host's preference.
					if (values.officeTheme !== 'dark') writeStatusBarTheme(values.officeTheme);
				}
				const initial = String(values.userInitials ?? '').trim();
				this.#update({
					displayName: String(values.userName ?? '').trim(),
					avatarColor: this.#profile.avatarColor,
					...(initial ? { initial } : {}),
				});
			},
			options,
		);
		this.#account?.addEventListener(
			'office-profile-change',
			(event) => this.#update((event as OfficeProfileChangeEvent).detail.profile),
			options,
		);
		return () => events.abort();
	}
	#update(profile: OfficeProfile): void {
		// Kept for this session even when storage is unavailable (private browsing).
		this.#profile = sanitizeOfficeProfile(profile);
		writeOfficeProfile(this.#profile);
		this.#sync();
		this.onProfile(this.profile);
	}
	#sync(): void {
		if (this.#account) this.#account.profile = this.#profile;
		this.#dialog.values = profileValues(this.#profile, this.theme.get());
	}
}
