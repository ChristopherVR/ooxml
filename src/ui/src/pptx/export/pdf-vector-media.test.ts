// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest';

import { shownVideoPoster } from './pdf-vector-media';

const POSTER = 'data:image/png;base64,iVBORw0KGgo=';

/** A stand-in for a mounted `<video>` with just the state the export reads. */
function video(
	state: {
		poster?: string;
		readyState?: number;
		paused?: boolean;
		played?: number;
		currentTime?: number;
	} = {},
) {
	const poster = 'poster' in state ? state.poster : POSTER;
	return {
		poster: poster ?? '',
		readyState: state.readyState ?? 4,
		paused: state.paused ?? true,
		played: { length: state.played ?? 0 } as TimeRanges,
		currentTime: state.currentTime ?? 0,
		getAttribute: (name: string) => (name === 'poster' ? (poster ?? null) : null),
	};
}

describe('the picture a stage video shows in the vector PDF', () => {
	it('is the poster for a paused background clip that has never played', () => {
		// A slide background clip: metadata and a first frame loaded, paused at the start. The
		// browser paints the poster, so the PDF must not draw the clip's (often black) first frame.
		expect(shownVideoPoster(video())).toBe(POSTER);
	});

	it('is the poster while no frame is decoded yet', () => {
		expect(shownVideoPoster(video({ readyState: 1, paused: false, played: 1 }))).toBe(POSTER);
	});

	it('is the current frame once the video has played or been seeked', () => {
		expect(shownVideoPoster(video({ played: 1, currentTime: 3 }))).toBeUndefined();
		expect(shownVideoPoster(video({ currentTime: 2 }))).toBeUndefined();
		expect(shownVideoPoster(video({ paused: false, played: 1 }))).toBeUndefined();
	});

	it('is the current frame when there is no poster', () => {
		expect(shownVideoPoster(video({ poster: undefined }))).toBeUndefined();
		expect(shownVideoPoster(video({ poster: undefined, readyState: 0 }))).toBeUndefined();
	});

	it('reads the poster of a real video element', () => {
		const el = document.createElement('video');
		el.setAttribute('poster', POSTER);
		expect(shownVideoPoster(el)).toBe(POSTER);
	});
});
