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
import { VISIO_OPTION_CATEGORIES, MAX_UNDO_LEVELS } from './viewer-options-categories';
import { isScreenTipStyle, type ScreenTipStyle } from './viewer-screen-tips';

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

/** The Options that change how the viewer behaves, remembered on this device. */
export interface VisioBehaviour {
	screenTips: ScreenTipStyle;
	undoLevels: number;
}
export const DEFAULT_BEHAVIOUR: VisioBehaviour = {
	screenTips: 'descriptions',
	undoLevels: MAX_UNDO_LEVELS,
};
const BEHAVIOUR_KEY = 'ooxml-visio-options';
const clampUndo = (value: unknown): number => {
	const levels = Math.round(Number(value));
	return Number.isFinite(levels)
		? Math.min(MAX_UNDO_LEVELS, Math.max(1, levels))
		: DEFAULT_BEHAVIOUR.undoLevels;
};
/** The behaviour Options saved on this device (storage may be unavailable or hold junk). */
export function readBehaviour(): VisioBehaviour {
	try {
		const stored = JSON.parse(globalThis.localStorage?.getItem(BEHAVIOUR_KEY) ?? '{}') as Record<
			string,
			unknown
		>;
		return {
			screenTips: isScreenTipStyle(stored.screenTips)
				? stored.screenTips
				: DEFAULT_BEHAVIOUR.screenTips,
			undoLevels: clampUndo(stored.undoLevels ?? DEFAULT_BEHAVIOUR.undoLevels),
		};
	} catch {
		return { ...DEFAULT_BEHAVIOUR };
	}
}
function writeBehaviour(behaviour: VisioBehaviour): void {
	try {
		globalThis.localStorage?.setItem(BEHAVIOUR_KEY, JSON.stringify(behaviour));
	} catch {
		// Kept for this session only.
	}
}

function profileValues(
	profile: OfficeProfile,
	officeTheme: OfficeThemeChoice,
	behaviour: VisioBehaviour,
): OfficeOptionValues {
	return {
		officeTheme,
		miniToolbar: false,
		livePreview: false,
		screenTips: behaviour.screenTips,
		undoLevels: behaviour.undoLevels,
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
	dialog.defaults = {
		userName: '',
		userInitials: '',
		officeTheme: 'neutral',
		...DEFAULT_BEHAVIOUR,
	};
	return dialog;
}

/**
 * File > Account and File > Options share one local profile. Edits in either are written to the
 * suite-wide profile store, mirrored into the other and reported through `onProfile`.
 */
export class ViewerProfile {
	#profile: OfficeProfile = readOfficeProfile();
	#behaviour: VisioBehaviour = readBehaviour();
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
		/** Applies the behaviour Options (ScreenTip style, undo depth), now and after each change. */
		private readonly onBehaviour: (behaviour: VisioBehaviour) => void = () => {},
	) {
		this.#dialog = root.querySelector<OptionsDialog>('office-ui-options-dialog')!;
		this.#account = root.querySelector<Account>('office-ui-account');
		this.#sync();
		this.onBehaviour({ ...this.#behaviour });
	}
	get profile(): OfficeProfile {
		return { ...this.#profile };
	}
	/** Open Visio Options, on General as Visio does. */
	showOptions(category = 'general'): void {
		this.#dialog.values = profileValues(this.#profile, this.theme.get(), this.#behaviour);
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
				const screenTips = isScreenTipStyle(values.screenTips)
					? values.screenTips
					: this.#behaviour.screenTips;
				const undoLevels = clampUndo(values.undoLevels ?? this.#behaviour.undoLevels);
				if (
					screenTips !== this.#behaviour.screenTips ||
					undoLevels !== this.#behaviour.undoLevels
				) {
					this.#behaviour = { screenTips, undoLevels };
					writeBehaviour(this.#behaviour);
					this.onBehaviour({ ...this.#behaviour });
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
		this.#dialog.values = profileValues(this.#profile, this.theme.get(), this.#behaviour);
	}
}
