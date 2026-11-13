import {
	BUILDING_BLOCK_OK,
	buildingBlockFailure,
	type BuildingBlockResult,
} from './building-block-actions';
import { describeFontEmbedding, type FontEmbeddingDescriptor } from './font-embedding';

/**
 * Framework-free font-embedding workflow: the File > Fonts panel's options for a host-owned shell.
 * The viewer can keep or strip the fonts a deck already carries; it cannot add a font that is not
 * embedded (a browser does not hand over the bytes of an installed face), so turning embedding on
 * for a deck that embeds nothing is refused with `unsupported`.
 */

export interface FontEmbeddingState {
	/** Saving keeps (true) or strips (false) the embedded fonts. */
	embedFonts: boolean;
	/** The embedded typefaces the loaded deck carries. */
	embeddedFamilies: string[];
	/** The typefaces the slides use. */
	usedFamilies: string[];
	/** The option can be turned on; false when the deck embeds nothing. */
	canEmbed: boolean;
	/** i18n key explaining why `canEmbed` is false. */
	unavailableReasonKey?: string;
}

export function describeFontEmbeddingState(input: {
	embedFonts: boolean;
	embeddedFamilies: readonly string[];
	usedFamilies: readonly string[];
	descriptor?: FontEmbeddingDescriptor;
}): FontEmbeddingState {
	const descriptor = input.descriptor ?? describeFontEmbedding(input.embeddedFamilies);
	return {
		embedFonts: input.embedFonts,
		embeddedFamilies: descriptor.embeddedFamilies,
		usedFamilies: [...input.usedFamilies],
		canEmbed: descriptor.interactive,
		...(descriptor.disabledReasonKey ? { unavailableReasonKey: descriptor.disabledReasonKey } : {}),
	};
}

export interface FontEmbeddingActions {
	state: FontEmbeddingState;
	/** Keep or strip the embedded fonts on save. */
	setEmbedFonts: (enabled: boolean) => BuildingBlockResult;
}

export interface FontEmbeddingActionsInput {
	state: FontEmbeddingState;
	refuse: () => BuildingBlockResult | undefined;
	/** The latest option value and whether embedding is possible. */
	latest: () => Pick<FontEmbeddingState, 'embedFonts' | 'canEmbed'>;
	apply: (enabled: boolean) => void;
	markDirty: () => void;
	complete: (result: BuildingBlockResult) => BuildingBlockResult;
}

export function buildFontEmbeddingActions(input: FontEmbeddingActionsInput): FontEmbeddingActions {
	return {
		state: input.state,
		setEmbedFonts: (enabled) => {
			const denied = input.refuse();
			if (denied) {
				return denied;
			}
			const latest = input.latest();
			if (enabled && !latest.canEmbed) {
				return buildingBlockFailure('unsupported');
			}
			if (latest.embedFonts !== enabled) {
				input.apply(enabled);
				input.markDirty();
			}
			return input.complete(BUILDING_BLOCK_OK);
		},
	};
}
