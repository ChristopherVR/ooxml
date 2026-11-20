/**
 * Slide-transition sound playback (`p:transition/p:sndAc/p:stSnd` and
 * `p:endSndAc`), as opposed to `slide-transition-sound.ts` (the AUTHORING
 * picker that lets the Transitions ribbon attach/clear a sound file).
 *
 * PowerPoint's Transitions ribbon > Sound dropdown attaches a sound to a
 * slide's transition: `p:stSnd` (with `@loop`, "Loop Until Next Sound") plays
 * a new sound the instant the transition starts; `p:endSnd` ("Stop Previous
 * Sound") silences whatever transition sound is already playing instead.
 *
 * {@link resolveTransitionSoundAction} is the pure decision (what a
 * transition's sound-action list means, given nothing but the transition
 * object); {@link applySlideTransitionSound} plays it through `deck-sound`, the
 * same one-sound-at-a-time channel as effect sounds, so the two constructs never
 * talk over each other and a linked transition sound waits for the viewer's
 * linked-content consent like every other deck sound.
 *
 * @module render/slide-transition-sound-playback
 */
import type { PptxSlideTransition } from 'ooxml-core/pptx';

import { playDeckSound, stopDeckSound } from './deck-sound';
import type { DeckSoundSources } from './deck-sound';

/**
 * What a slide transition's sound-action list means at the moment the
 * transition starts.
 *
 * `'play'` and `'stop'` are mutually exclusive on a single transition (core's
 * `PptxSlideTransition` comment on `stopSound` says as much: `p:stSnd` and
 * `p:endSnd` are alternatives within `p:sndAc`), so this is a plain union
 * rather than two independent booleans.
 */
export type TransitionSoundAction =
	| { kind: 'play'; soundPath: string; loop: boolean }
	| { kind: 'stop' }
	| { kind: 'none' };

/**
 * Decide what `transition`'s sound action means, without resolving the sound
 * path to a playable URL (see {@link applySlideTransitionSound}).
 */
export function resolveTransitionSoundAction(
	transition: PptxSlideTransition | undefined,
): TransitionSoundAction {
	if (!transition) {
		return { kind: 'none' };
	}
	if (transition.soundPath) {
		return { kind: 'play', soundPath: transition.soundPath, loop: transition.soundLoop === true };
	}
	if (transition.stopSound) {
		return { kind: 'stop' };
	}
	return { kind: 'none' };
}

/**
 * Play or stop `transition`'s sound the instant the transition starts. The
 * sound is an archive path the viewer's load pipeline put in
 * `sources.mediaUrls`, or a linked web address played only once
 * `sources.linkedContent` allows it.
 */
export function applySlideTransitionSound(
	transition: PptxSlideTransition | undefined,
	sources: DeckSoundSources,
): void {
	const action = resolveTransitionSoundAction(transition);
	if (action.kind === 'stop') {
		stopDeckSound();
	} else if (action.kind === 'play') {
		playDeckSound(action.soundPath, sources, { loop: action.loop });
	}
}
