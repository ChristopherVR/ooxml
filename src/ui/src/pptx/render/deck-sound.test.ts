import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { playDeckSound, resolveDeckSound, stopDeckSound } from './deck-sound';

interface FakeAudio {
	src: string;
	currentTime: number;
	loop: boolean;
	play: ReturnType<typeof vi.fn>;
	pause: ReturnType<typeof vi.fn>;
}

const LINKED = 'https://linked-content.test/click.wav';
const BLOB = 'blob:https://app.test/1';
const DATA = 'data:audio/wav;base64,AAAA';
const MEDIA = new Map([
	['ppt/media/audio1.wav', BLOB],
	['ppt/media/linked.wav', LINKED],
]);

let created: FakeAudio[] = [];

beforeEach(() => {
	created = [];
	vi.stubGlobal(
		'Audio',
		// oxlint-disable-next-line prefer-arrow-callback -- `new Audio()` needs a function
		vi.fn(function (src: string) {
			const audio: FakeAudio = {
				src,
				currentTime: 0,
				loop: false,
				play: vi.fn<() => Promise<void>>().mockResolvedValue(undefined),
				pause: vi.fn<() => void>(),
			};
			created.push(audio);
			return audio;
		}),
	);
});

afterEach(() => {
	stopDeckSound();
	vi.unstubAllGlobals();
});

describe('resolveDeckSound', () => {
	it('withholds a linked sound until the viewer allows linked content', () => {
		expect(resolveDeckSound(LINKED, { linkedContent: false })).toBeUndefined();
		expect(resolveDeckSound(LINKED, { linkedContent: true })).toBe(LINKED);
	});

	it('plays an embedded sound through the media map, and data or blob URLs as they are', () => {
		expect(
			resolveDeckSound('ppt/media/audio1.wav', { mediaUrls: MEDIA, linkedContent: false }),
		).toBe(BLOB);
		expect(resolveDeckSound(DATA, { linkedContent: false })).toBe(DATA);
		expect(resolveDeckSound(BLOB, { linkedContent: false })).toBe(BLOB);
	});

	it('never hands the browser an archive path or a web address the media map holds', () => {
		expect(resolveDeckSound('ppt/media/missing.wav', { linkedContent: true })).toBeUndefined();
		expect(
			resolveDeckSound('ppt/media/linked.wav', { mediaUrls: MEDIA, linkedContent: false }),
		).toBeUndefined();
		expect(
			resolveDeckSound('ppt/media/linked.wav', { mediaUrls: MEDIA, linkedContent: true }),
		).toBe(LINKED);
		expect(resolveDeckSound(undefined, { linkedContent: true })).toBeUndefined();
	});
});

describe('playDeckSound', () => {
	it('fetches nothing for a linked sound before consent, and plays it after', () => {
		playDeckSound(LINKED, { linkedContent: false });
		expect(Audio).not.toHaveBeenCalled();
		playDeckSound(LINKED, { linkedContent: true });
		expect(Audio).toHaveBeenCalledWith(LINKED);
		expect(created[0]!.play).toHaveBeenCalledOnce();
	});

	it('replaces the sound playing, and loops when asked', () => {
		playDeckSound(DATA, { linkedContent: false });
		playDeckSound(BLOB, { linkedContent: false }, { loop: true });
		expect(created[0]!.pause).toHaveBeenCalledOnce();
		expect(created[0]!.loop).toBe(false);
		expect(created[1]!.loop).toBe(true);
	});

	it('plays an overlapping sound alongside the current one', () => {
		playDeckSound(DATA, { linkedContent: false }, { loop: true });
		playDeckSound(BLOB, { linkedContent: false }, { overlap: true });
		expect(created[0]!.pause).not.toHaveBeenCalled();
		stopDeckSound();
		expect(created[0]!.pause).toHaveBeenCalledOnce();
		expect(created[1]!.pause).not.toHaveBeenCalled();
	});

	it('stops once, and ignores autoplay refusals', () => {
		vi.mocked(Audio).mockImplementationOnce(function (src?: string) {
			const audio = {
				src,
				currentTime: 3,
				loop: false,
				play: vi.fn<() => Promise<void>>().mockRejectedValue(new Error('blocked')),
				pause: vi.fn<() => void>(),
			};
			created.push(audio as FakeAudio);
			return audio as unknown as HTMLAudioElement;
		});
		expect(() => playDeckSound(DATA, { linkedContent: false })).not.toThrow();
		stopDeckSound();
		stopDeckSound();
		expect(created[0]!.pause).toHaveBeenCalledOnce();
		expect(created[0]!.currentTime).toBe(0);
	});
});
