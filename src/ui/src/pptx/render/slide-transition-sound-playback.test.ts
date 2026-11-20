import type { PptxSlideTransition } from 'ooxml-core/pptx';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { stopDeckSound } from './deck-sound';
import {
	applySlideTransitionSound,
	resolveTransitionSoundAction,
} from './slide-transition-sound-playback';

describe('resolveTransitionSoundAction', () => {
	it('returns none for an undefined transition', () => {
		expect(resolveTransitionSoundAction(undefined)).toStrictEqual({ kind: 'none' });
	});

	it('plays p:stSnd with loop=false by default', () => {
		const transition = { type: 'fade', soundPath: 'ppt/media/media3.wav' } as PptxSlideTransition;
		expect(resolveTransitionSoundAction(transition)).toStrictEqual({
			kind: 'play',
			soundPath: 'ppt/media/media3.wav',
			loop: false,
		});
	});

	it('honors @loop ("Loop Until Next Sound")', () => {
		const transition = {
			type: 'fade',
			soundPath: 'ppt/media/media3.wav',
			soundLoop: true,
		} as PptxSlideTransition;
		expect(resolveTransitionSoundAction(transition)).toStrictEqual({
			kind: 'play',
			soundPath: 'ppt/media/media3.wav',
			loop: true,
		});
	});

	it('stops the current sound for p:endSndAc (stopSound)', () => {
		const transition = { type: 'fade', stopSound: true } as PptxSlideTransition;
		expect(resolveTransitionSoundAction(transition)).toStrictEqual({ kind: 'stop' });
	});

	it('prefers play over stop when a transition somehow authors both', () => {
		const transition = {
			type: 'fade',
			soundPath: 'ppt/media/media3.wav',
			stopSound: true,
		} as PptxSlideTransition;
		expect(resolveTransitionSoundAction(transition).kind).toBe('play');
	});

	it('returns none for a transition with neither a sound nor a stop flag', () => {
		expect(resolveTransitionSoundAction({ type: 'fade' } as PptxSlideTransition)).toStrictEqual({
			kind: 'none',
		});
	});
});

describe('applySlideTransitionSound', () => {
	let created: string[] = [];
	let paused = 0;

	beforeEach(() => {
		created = [];
		paused = 0;
		vi.stubGlobal(
			'Audio',
			// oxlint-disable-next-line prefer-arrow-callback -- `new Audio()` needs a function
			vi.fn(function (src: string) {
				created.push(src);
				return {
					loop: false,
					currentTime: 0,
					play: () => Promise.resolve(),
					pause: () => {
						paused += 1;
					},
				};
			}),
		);
	});

	afterEach(() => {
		stopDeckSound();
		vi.unstubAllGlobals();
	});

	it('plays an embedded sound from the media map', () => {
		const transition = { type: 'fade', soundPath: 'a.wav', soundLoop: true } as PptxSlideTransition;
		applySlideTransitionSound(transition, {
			mediaUrls: new Map([['a.wav', 'blob:xyz']]),
			linkedContent: false,
		});
		expect(created).toStrictEqual(['blob:xyz']);
	});

	it('withholds a linked transition sound until the viewer allows linked content', () => {
		const transition = {
			type: 'fade',
			soundPath: 'https://linked-content.test/whoosh.wav',
		} as PptxSlideTransition;
		applySlideTransitionSound(transition, { linkedContent: false });
		expect(created).toStrictEqual([]);
		applySlideTransitionSound(transition, { linkedContent: true });
		expect(created).toStrictEqual(['https://linked-content.test/whoosh.wav']);
	});

	it('plays nothing for a sound the media map does not hold', () => {
		applySlideTransitionSound({ type: 'fade', soundPath: 'missing.wav' } as PptxSlideTransition, {
			linkedContent: true,
		});
		expect(created).toStrictEqual([]);
	});

	it('stops the playing sound for p:endSndAc', () => {
		applySlideTransitionSound({ type: 'fade', soundPath: 'data:audio/wav;base64,AA' } as never, {
			linkedContent: false,
		});
		applySlideTransitionSound({ type: 'fade', stopSound: true } as PptxSlideTransition, {
			linkedContent: false,
		});
		expect(paused).toBe(1);
	});
});
