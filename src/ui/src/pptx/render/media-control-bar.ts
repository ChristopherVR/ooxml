/**
 * The PowerPoint-style media control bar of the authoring canvas.
 *
 * PowerPoint's Normal view paints no transport on a video or audio shape at
 * rest (and a static export shows none); a bar with play/pause, a scrubber and
 * a volume toggle appears BELOW the shape while it is selected or hovered. The
 * bar is a sibling of the media element, not Chrome's own `controls`, so a
 * click on the picture still selects or moves the shape and a click on the bar
 * plays the clip without ever reaching the select/move gesture.
 *
 * One imperative helper serves all five bindings: each hands it the element's
 * host box and the `<video>`/`<audio>` it already renders, plus whether the
 * element is selected, and disposes it with the component.
 */
import { MEDIA_CHROME_ATTRIBUTE, type MediaSurface } from './media-playback';

/**
 * Whether the authoring canvas offers the control bar: only on the editable
 * stage. A show and a still of a slide never carry one, and the reading view
 * keeps the browser's native transport instead.
 */
export function mediaControlBarEnabled(surface: MediaSurface & { reading?: boolean }): boolean {
	return !surface.presenting && !surface.preview && surface.reading !== true;
}

/** Whether the bar shows right now: selected or hovered, never at rest. */
export function mediaControlBarVisible(state: { selected: boolean; hovered: boolean }): boolean {
	return state.selected || state.hovered;
}

export interface MediaControlBar {
	/** The bar element (hidden at rest). */
	readonly element: HTMLElement;
	setSelected(selected: boolean): void;
	dispose(): void;
}

export interface MediaControlBarOptions {
	/** Translator; uses `pptx.media.play`, `pptx.media.pause` and `pptx.media.volume`. */
	t: (key: string) => string;
	selected?: boolean;
}

const BAR_STYLE =
	'position:absolute;left:0;top:100%;min-width:168px;width:100%;box-sizing:border-box;' +
	'padding-top:4px;z-index:2147483000;pointer-events:auto;';
const BODY_STYLE =
	'display:flex;align-items:center;gap:6px;height:28px;padding:0 8px;box-sizing:border-box;' +
	'background:#262626;color:#fff;border-radius:4px;font:12px sans-serif;';
const BUTTON_STYLE =
	'all:unset;cursor:pointer;width:20px;text-align:center;line-height:20px;color:inherit;';

export function attachMediaControlBar(
	host: HTMLElement,
	media: HTMLMediaElement,
	options: MediaControlBarOptions,
): MediaControlBar {
	const doc = host.ownerDocument;
	const bar = doc.createElement('div');
	bar.setAttribute(MEDIA_CHROME_ATTRIBUTE, 'control-bar');
	bar.setAttribute('data-pptx-media-bar', '');
	bar.setAttribute('style', BAR_STYLE);
	const body = doc.createElement('div');
	body.setAttribute('style', BODY_STYLE);
	const play = doc.createElement('button');
	play.type = 'button';
	play.setAttribute('style', BUTTON_STYLE);
	const seek = doc.createElement('input');
	seek.type = 'range';
	seek.min = '0';
	seek.max = '1000';
	seek.value = '0';
	seek.setAttribute('style', 'flex:1;min-width:0;margin:0;cursor:pointer;');
	const mute = doc.createElement('button');
	mute.type = 'button';
	mute.setAttribute('style', BUTTON_STYLE);
	mute.setAttribute('aria-label', options.t('pptx.media.volume'));
	body.append(play, seek, mute);
	bar.append(body);
	host.append(bar);

	let selected = options.selected === true;
	let hovered = false;
	const refresh = (): void => {
		bar.hidden = !mediaControlBarVisible({ selected, hovered });
		bar.style.display = bar.hidden ? 'none' : '';
		const label = options.t(media.paused ? 'pptx.media.play' : 'pptx.media.pause');
		play.setAttribute('aria-label', label);
		play.title = label;
		play.textContent = media.paused ? '▶' : '❚❚';
		mute.textContent = media.muted || media.volume === 0 ? '\u{1F507}' : '\u{1F50A}';
		const duration = Number.isFinite(media.duration) && media.duration > 0 ? media.duration : 0;
		if (doc.activeElement !== seek) {
			seek.value = duration ? String(Math.round((media.currentTime / duration) * 1000)) : '0';
		}
	};

	// Hover is decided from the pointer position against the shape's box and the
	// bar's box, not from pointerenter on the host: the editable canvases hit-test
	// by coordinates and keep the media element itself `pointer-events: none`, so
	// the host never receives enter/leave there. The padding-top above closes the
	// gap between the shape and the bar.
	const within = (rect: DOMRect, event: MouseEvent): boolean =>
		event.clientX >= rect.left &&
		event.clientX <= rect.right &&
		event.clientY >= rect.top &&
		event.clientY <= rect.bottom;
	const setHovered = (next: boolean): void => {
		if (next !== hovered) {
			hovered = next;
			refresh();
		}
	};
	const onMove = (event: Event): void => {
		const pointer = event as MouseEvent;
		setHovered(
			within(host.getBoundingClientRect(), pointer) ||
				(!bar.hidden && within(bar.getBoundingClientRect(), pointer)),
		);
	};
	const onLeaveDocument = (): void => setHovered(false);
	// A press on the bar is the bar's own: it must not select, move or step.
	const swallow = (event: Event): void => event.stopPropagation();
	const togglePlay = (): void => {
		if (media.paused) {
			void media.play().catch(() => undefined);
		} else {
			media.pause();
		}
	};
	const toggleMute = (): void => {
		media.muted = !media.muted;
	};
	const scrub = (): void => {
		const duration = Number.isFinite(media.duration) ? media.duration : 0;
		if (duration > 0) {
			media.currentTime = (Number(seek.value) / 1000) * duration;
		}
	};
	const mediaEvents = ['play', 'pause', 'timeupdate', 'loadedmetadata', 'volumechange', 'ended'];
	doc.addEventListener('pointermove', onMove, true);
	doc.documentElement.addEventListener('pointerleave', onLeaveDocument);
	for (const type of ['pointerdown', 'mousedown', 'click', 'dblclick', 'contextmenu']) {
		bar.addEventListener(type, swallow);
	}
	play.addEventListener('click', togglePlay);
	mute.addEventListener('click', toggleMute);
	seek.addEventListener('input', scrub);
	for (const type of mediaEvents) {
		media.addEventListener(type, refresh);
	}
	refresh();

	return {
		element: bar,
		setSelected(next) {
			selected = next;
			refresh();
		},
		dispose() {
			doc.removeEventListener('pointermove', onMove, true);
			doc.documentElement.removeEventListener('pointerleave', onLeaveDocument);
			for (const type of mediaEvents) {
				media.removeEventListener(type, refresh);
			}
			bar.remove();
		},
	};
}
