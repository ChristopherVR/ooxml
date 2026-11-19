/**
 * The one way a viewer plays a sound from a deck: animation, transition and
 * hyperlink sounds in the slide show, and the Sound pickers' previews.
 *
 * A deck names a sound by an archive path (`ppt/media/media1.wav`, which the
 * viewer's load pipeline has turned into a blob URL in its media map) or, when
 * the sound is linked instead of embedded, by a web address. A linked sound is
 * linked content like a linked picture: it may be fetched only once this viewer
 * allows linked content (`resolveExternalContentView`). {@link playDeckSound}
 * takes that consent with every call, so no sound reaches the browser without
 * it, and nothing else in the bindings creates an `Audio` (the viewers' lint
 * forbids the global outside this module).
 *
 * One sound plays at a time, as in PowerPoint: a new sound replaces the one
 * playing, unless it is played with `overlap`.
 *
 * @module render/deck-sound
 */
import { isWebAddress } from './options/linked-content';

/** Where a viewer may take a deck's sounds from, read when a sound plays. */
export interface DeckSoundSources {
	/** The viewer's archive path to blob URL map, filled by its load pipeline. */
	readonly mediaUrls?: ReadonlyMap<string, string> | undefined;
	/** Whether this viewer allows the deck's linked content now. */
	readonly linkedContent: boolean;
}

/**
 * Sources for a sound that is not the deck's: the Sound pickers' built-in sounds
 * (data URLs). Nothing linked plays through it.
 */
export const BUILT_IN_SOUNDS: DeckSoundSources = Object.freeze({ linkedContent: false });

/** How {@link playDeckSound} plays a sound. */
export interface PlayDeckSoundOptions {
	/** Repeat until stopped or replaced ("Loop Until Next Sound"). */
	readonly loop?: boolean;
	/** Play alongside the current sound instead of replacing it (a hyperlink's sound). */
	readonly overlap?: boolean;
}

function isLocalUrl(value: string): boolean {
	return /^(?:data|blob):/iu.test(value);
}

/**
 * The address the browser may play for `source`: a data or blob URL as it is,
 * an archive path through `mediaUrls`, and a web address only when
 * `linkedContent` allows it. `undefined` when there is nothing it may play.
 */
export function resolveDeckSound(
	source: string | undefined,
	sources: DeckSoundSources,
): string | undefined {
	if (!source) {
		return undefined;
	}
	if (isWebAddress(source)) {
		return sources.linkedContent ? source : undefined;
	}
	if (isLocalUrl(source)) {
		return source;
	}
	const mapped = sources.mediaUrls?.get(source);
	return mapped === undefined
		? undefined
		: resolveDeckSound(mapped, { linkedContent: sources.linkedContent });
}

let current: HTMLAudioElement | undefined;

/**
 * Play the deck sound `source` names, as {@link resolveDeckSound} allows it.
 * Silently ignores a sound it may not play, autoplay refusals and runtimes
 * without audio.
 */
export function playDeckSound(
	source: string | undefined,
	sources: DeckSoundSources,
	options: PlayDeckSoundOptions = {},
): void {
	const url = resolveDeckSound(source, sources);
	if (url === undefined || typeof Audio === 'undefined') {
		return;
	}
	if (!options.overlap) {
		stopDeckSound();
	}
	const audio = new Audio(url);
	audio.loop = options.loop === true;
	if (!options.overlap) {
		current = audio;
	}
	audio.play().catch(() => {
		/* autoplay restrictions and unplayable sources are not errors here */
	});
}

/** Stop the sound {@link playDeckSound} last started without `overlap`, if any. */
export function stopDeckSound(): void {
	if (current) {
		current.pause();
		current.currentTime = 0;
		current = undefined;
	}
}
