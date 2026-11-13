import { describeBuildingBlockActions } from './building-block-actions';
import type {
	BuildingBlockActionState,
	BuildingBlockAvailabilityInput,
} from './building-block-actions';
import { VIEWER_SHORTCUT_REFERENCE, groupShortcutReference } from './viewer-preferences';
import type { ShortcutReferenceGroup } from './viewer-preferences';

/**
 * Framework-free parts of the slide show dialogs a host renders itself: Set Up Slide Show, Custom
 * Shows, the Broadcast dialog and the keyboard shortcut help. The settings and show edits are in
 * `building-block-show-settings` and `building-block-custom-shows`; this file holds the request
 * kinds, which of the four can run now, the shortcut catalogue as data, and the honest answer for
 * Broadcast.
 */

/** What a toolbar control or shortcut asked the host to show. */
export type ShowDialogRequest =
	| { kind: 'setUpSlideShow' }
	| { kind: 'customShows' }
	/** The shortcut help was opened (`open`) or closed; the host shows or hides its own. */
	| { kind: 'shortcuts'; open: boolean };

export type ShowDialogActionId = 'setUpSlideShow' | 'customShows' | 'broadcast' | 'shortcuts';

/**
 * Which of the four can run now. Set Up and Custom Shows write the deck, so they follow the
 * same edit gate as the other workflows. The shortcut help is read-only data and always
 * available. Broadcast needs a collaboration service the building blocks do not run, so it is
 * never available here (see {@link describeBroadcastSupport}).
 */
export function describeShowDialogActions(
	input: BuildingBlockAvailabilityInput,
): Record<ShowDialogActionId, BuildingBlockActionState> {
	// `replace` has no condition of its own: its state is the shared edit gate.
	const gate = describeBuildingBlockActions(input).replace;
	const edit = (): BuildingBlockActionState =>
		gate.available
			? { available: true, needsHostUi: true }
			: { available: false, reason: gate.reason ?? 'not-ready', needsHostUi: true };
	return {
		setUpSlideShow: edit(),
		customShows: edit(),
		broadcast: { available: false, reason: 'unsupported', needsHostUi: true },
		shortcuts: { available: true, needsHostUi: true },
	};
}

/**
 * Broadcast streams the presenter's slide to viewers through a collaboration service (a Yjs
 * WebSocket server or peer-to-peer signaling) and enters a presenter session. The building
 * blocks neither run nor own one, so they cannot start or stop a broadcast. A host that has its
 * own service passes `collaboration` (role `owner`) to the hook and presents with its own
 * controls.
 */
export interface BroadcastSupport {
	supported: false;
	reason: 'unsupported';
	/** Why: no service is run by the building blocks. */
	requires: 'collaboration-service';
}

export function describeBroadcastSupport(): BroadcastSupport {
	return { supported: false, reason: 'unsupported', requires: 'collaboration-service' };
}

/** One entry of the keyboard shortcut help, translated. */
export interface ShortcutCatalogueEntry {
	/** Stable id, from the translation key (for example `toolbar.undo`). */
	id: string;
	/** The translation key of the description, for a host with its own dictionary. */
	descriptionKey: string;
	/** The description in the viewer's current language. */
	description: string;
	/** The keys as the viewer's help shows them (`Ctrl/Cmd` stands for either). */
	keys: string;
	/** `keys` split into its alternatives (`Ctrl/Cmd+Y`, ...). */
	alternatives: string[];
	group: ShortcutReferenceGroup;
	/** The group heading in the viewer's current language. */
	groupLabel: string;
}

/** The shortcut help as data, in help order, descriptions and headings through `translate`. */
export function describeShortcutCatalogue(
	translate: (key: string) => string,
): ShortcutCatalogueEntry[] {
	return groupShortcutReference(VIEWER_SHORTCUT_REFERENCE).flatMap((bucket) =>
		bucket.shortcuts.map((item) => ({
			id: item.actionKey.replace(/^pptx\./, '').replace(/^shortcuts\.action\./, ''),
			descriptionKey: item.actionKey,
			description: translate(item.actionKey),
			keys: item.shortcut,
			alternatives: item.shortcut.split(' or ').map((part) => part.trim()),
			group: item.group,
			groupLabel: translate(bucket.labelKey),
		})),
	);
}
