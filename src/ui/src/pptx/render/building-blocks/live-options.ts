/**
 * Which building-block options apply to a mounted canvas without remounting it, and how a binding
 * whose engine reads an option only once applies a change. React, Vue, Svelte and Angular follow
 * every live option through their own reactivity; the vanilla engine is imperative, so its blocks
 * run {@link applyBuildingBlockLiveOptions} from `update()`.
 *
 * @module render/building-blocks/live-options
 */
import type { CollaborationInput } from '../collaboration-options';
import { resolveCollaborationConfig } from '../collaboration-options';
import { resolveGoogleFontsCss2Base } from '../google-webfonts';
import type { GoogleFontsOption } from '../google-webfonts';
import { hostFontDownloadLock } from '../options/font-download-policy';

/**
 * The options every binding's building blocks apply to a mounted canvas at once, keeping the
 * open document (its edits, selection and undo history). `limits` and `initialSlide` apply to
 * the next load; the rest of a binding's viewer options are documented per binding.
 */
export const BUILDING_BLOCK_LIVE_OPTIONS = [
	'content',
	'canEdit',
	'onRequest',
	'fileLimits',
	'hiddenActions',
	'listConversionScope',
	'collaboration',
	'googleFonts',
	'allowExternalContent',
] as const;

export type BuildingBlockLiveOption = (typeof BUILDING_BLOCK_LIVE_OPTIONS)[number];

/** The options {@link applyBuildingBlockLiveOptions} compares. */
export interface BuildingBlockLiveOptionValues {
	collaboration?: CollaborationInput | null | undefined;
	googleFonts?: GoogleFontsOption | undefined;
	allowExternalContent?: boolean | undefined;
}

/** What an imperative engine does when a live option changed. */
export interface BuildingBlockLiveOptionPort {
	/** Join the room (any running session is left first). */
	startCollaboration(input: CollaborationInput): unknown;
	stopCollaboration(): void;
	/** Lock Trust Center > "Download fonts" to the host policy and refresh the deck's webfonts. */
	setGoogleFonts(option: GoogleFontsOption | undefined): void;
	/** Lock Trust Center > "Allow external content" (unset: the user decides). */
	setAllowExternalContent(allow: boolean | undefined): void;
}

/**
 * The collaboration a change of the option asks for: `start` with the new input, `stop`, or
 * nothing. Compared by the resolved config's identity (an options object re-created with the
 * same `collaboration` value does not rejoin), as every binding's viewer compares it.
 */
export function collaborationOptionChange(
	previous: CollaborationInput | null | undefined,
	next: CollaborationInput | null | undefined,
): { kind: 'start'; input: CollaborationInput } | { kind: 'stop' } | null {
	const before = resolveCollaborationConfig(previous);
	const after = resolveCollaborationConfig(next);
	if (before === after) {
		return null;
	}
	return next && after ? { kind: 'start', input: next } : before ? { kind: 'stop' } : null;
}

/** True when the host `googleFonts` policy changed what fonts load (the lock or the mirror). */
export function googleFontsOptionChanged(
	previous: GoogleFontsOption | undefined,
	next: GoogleFontsOption | undefined,
): boolean {
	return (
		hostFontDownloadLock(previous) !== hostFontDownloadLock(next) ||
		resolveGoogleFontsCss2Base(previous) !== resolveGoogleFontsCss2Base(next)
	);
}

/** Apply the live options that changed between `previous` and `next` to an imperative engine. */
export function applyBuildingBlockLiveOptions(
	previous: BuildingBlockLiveOptionValues,
	next: BuildingBlockLiveOptionValues,
	port: BuildingBlockLiveOptionPort,
): void {
	if (next.allowExternalContent !== previous.allowExternalContent) {
		port.setAllowExternalContent(next.allowExternalContent);
	}
	if (googleFontsOptionChanged(previous.googleFonts, next.googleFonts)) {
		port.setGoogleFonts(next.googleFonts);
	}
	const collaboration = collaborationOptionChange(previous.collaboration, next.collaboration);
	if (collaboration?.kind === 'start') {
		void port.startCollaboration(collaboration.input);
	} else if (collaboration?.kind === 'stop') {
		port.stopCollaboration();
	}
}
