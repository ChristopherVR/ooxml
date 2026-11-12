import { classifyMediaType } from './media-file-type';
import type { MediaFileType } from './media-file-type';

/**
 * Framework-free decisions behind the host-owned workflows of a custom shell built on the
 * building-block API (`useViewerBuildingBlocks`): which actions can run right now and why not,
 * and whether a file a host's own picker returned may be inserted. The binding only maps these
 * onto its state and operations. Messages are deliberately codes, not English text: the host
 * renders the feedback in its own language.
 */

/** Why a host-driven operation did not run. */
export type BuildingBlockFailureCode =
	/** The deck is still loading, failed to load, or has no active slide. */
	| 'not-ready'
	/** The viewer is not editable (view-only, a mode other than edit, or no edit permission). */
	| 'read-only'
	/** The deck is read-only recommended / in Protected View, so inserting is blocked. */
	| 'protected'
	| 'no-slide'
	| 'no-selection'
	| 'no-file'
	| 'empty-file'
	| 'unsupported-type'
	| 'file-too-large'
	/** The file could not be decoded (a corrupt or unreadable image). */
	| 'decode-failed'
	/** The slide, deck or selection changed while the operation was in flight. */
	| 'stale-target'
	/** A link target that is unsafe (script, data or similar scheme). */
	| 'unsafe-url'
	/** A link target with nothing to link to. */
	| 'invalid-target'
	| 'no-hyperlink'
	| 'no-equation'
	/** Find / Replace has no (current) match to act on. */
	| 'no-match'
	/** The layout is not one of the active slide's master's layouts. */
	| 'unknown-layout'
	| 'no-layout';

export type BuildingBlockResult = { ok: true } | { ok: false; code: BuildingBlockFailureCode };

export const BUILDING_BLOCK_OK: BuildingBlockResult = { ok: true };

export function buildingBlockFailure(code: BuildingBlockFailureCode): BuildingBlockResult {
	return { ok: false, code };
}

/** Default ceilings for a host-picked file (it is inlined as a data URL). */
export const HOST_IMAGE_MAX_BYTES = 25 * 1024 * 1024;
export const HOST_MEDIA_MAX_BYTES = 100 * 1024 * 1024;

export interface HostFileLimits {
	/** Largest accepted image, in bytes. Defaults to {@link HOST_IMAGE_MAX_BYTES}. */
	imageBytes?: number;
	/** Largest accepted audio or video file, in bytes. Defaults to {@link HOST_MEDIA_MAX_BYTES}. */
	mediaBytes?: number;
}

interface FileFacts {
	type: string;
	size: number;
}

/** Whether a host-picked image may be inserted (type, emptiness, size). */
export function validateHostImageFile(
	file: FileFacts | null | undefined,
	maxBytes: number = HOST_IMAGE_MAX_BYTES,
): BuildingBlockResult {
	if (!file) {
		return buildingBlockFailure('no-file');
	}
	if (!file.type.startsWith('image/')) {
		return buildingBlockFailure('unsupported-type');
	}
	if (file.size === 0) {
		return buildingBlockFailure('empty-file');
	}
	return file.size > maxBytes ? buildingBlockFailure('file-too-large') : BUILDING_BLOCK_OK;
}

/** Whether a host-picked audio or video file may be inserted, and which kind it is. */
export function validateHostMediaFile(
	file: FileFacts | null | undefined,
	maxBytes: number = HOST_MEDIA_MAX_BYTES,
): BuildingBlockResult & { mediaType?: MediaFileType } {
	if (!file) {
		return buildingBlockFailure('no-file');
	}
	const mediaType = classifyMediaType(file.type);
	if (!mediaType) {
		return buildingBlockFailure('unsupported-type');
	}
	if (file.size === 0) {
		return buildingBlockFailure('empty-file');
	}
	return file.size > maxBytes ? buildingBlockFailure('file-too-large') : { ok: true, mediaType };
}

export type BuildingBlockActionId =
	| 'insertImage'
	| 'insertMedia'
	| 'hyperlink'
	| 'removeHyperlink'
	| 'insertSmartArt'
	| 'insertEquation'
	| 'editEquation'
	| 'replace'
	| 'applyLayout'
	| 'resetLayout';

export interface BuildingBlockActionState {
	/** The operation would run now; when false, `reason` says why. */
	available: boolean;
	reason?: BuildingBlockFailureCode;
	/**
	 * The toolbar control for this action only raises a request: the host must render the picker or
	 * dialog (the building blocks mount none) and call the matching operation. False for actions
	 * that run straight from the toolbar or an operation call.
	 */
	needsHostUi: boolean;
}

export interface BuildingBlockAvailabilityInput {
	/** The deck is loaded without error. */
	ready: boolean;
	hasActiveSlide: boolean;
	canEdit: boolean;
	/** The viewer mode is `edit`. */
	editing: boolean;
	/** False while read-only recommended or in Protected View. */
	canInsert: boolean;
	hasSelection: boolean;
	selectionHasHyperlink: boolean;
	selectionHasEquation: boolean;
	/** The layouts the active slide may switch to (its master's). */
	layoutCount: number;
	/** The active slide records a layout, so Reset has one to return to. */
	hasCurrentLayout: boolean;
	/** The deck handler that applies a layout is available. */
	hasHandler: boolean;
}

/**
 * Toolbar controls whose dialog or pane the building blocks never mount: they set internal state
 * and nothing appears. A custom shell hides them (`hiddenActions`) or supplies its own UI.
 */
export const BUILDING_BLOCK_UNMOUNTED_DIALOG_ACTIONS = [
	'onOpenChartDialog',
	'onOpenPasswordProtection',
	'onOpenDocumentProperties',
	'onOpenFontEmbedding',
	'onOpenDigitalSignatures',
	'onOpenSetUpSlideShow',
	'onOpenCustomShows',
	'onOpenBroadcastDialog',
	'onToggleShortcuts',
] as const;

export function describeBuildingBlockActions(
	input: BuildingBlockAvailabilityInput,
): Record<BuildingBlockActionId, BuildingBlockActionState> {
	const gate = (): BuildingBlockFailureCode | undefined => {
		if (!input.ready) {
			return 'not-ready';
		}
		if (!input.canEdit || !input.editing) {
			return 'read-only';
		}
		return input.canInsert ? undefined : 'protected';
	};
	const state = (
		needsHostUi: boolean,
		...conditions: Array<BuildingBlockFailureCode | undefined>
	): BuildingBlockActionState => {
		const reason = [gate(), ...conditions].find((code) => code !== undefined);
		return reason === undefined
			? { available: true, needsHostUi }
			: { available: false, reason, needsHostUi };
	};
	const slide = input.hasActiveSlide ? undefined : 'no-slide';
	const selection = input.hasSelection ? undefined : 'no-selection';
	return {
		insertImage: state(true, slide),
		insertMedia: state(true, slide),
		hyperlink: state(true, selection),
		removeHyperlink: state(
			false,
			selection,
			input.selectionHasHyperlink ? undefined : 'no-hyperlink',
		),
		insertSmartArt: state(true, slide),
		insertEquation: state(true, slide),
		editEquation: state(true, selection, input.selectionHasEquation ? undefined : 'no-equation'),
		replace: state(false),
		applyLayout: state(
			false,
			slide,
			input.hasHandler && input.layoutCount > 0 ? undefined : 'no-layout',
		),
		resetLayout: state(
			false,
			slide,
			input.hasHandler && input.hasCurrentLayout ? undefined : 'no-layout',
		),
	};
}
