import { withTokens } from '../../base';
import { modalActiveElement } from '../../dialog/focus';
import { formatMediaTime } from '../render/media-trim-timeline';
import type { InsertText } from './ribbon-graphic-insert';
import css from './ribbon-recording-dock.css?raw';

/** Active capture controls stay available without keeping the area picker modal. */
export function showRecordingDock(
	host: HTMLElement,
	recorder: MediaRecorder,
	text: InsertText,
	discard: () => void,
): () => void {
	const doc = host.ownerDocument;
	const view = doc.defaultView;
	const listeners = new AbortController();
	const { signal } = listeners;
	const canHover = view?.matchMedia('(hover: hover)').matches === true;
	const dock = doc.createElement('div');
	dock.className = 'pptx-recording-dock';
	dock.setAttribute('role', 'group');
	dock.setAttribute('aria-label', text('pptx.insert.screenRecording', 'Screen Recording'));
	dock.setAttribute('popover', 'manual');
	const style = doc.createElement('style');
	style.textContent = withTokens(css);
	function button(caption: string, action: () => void): HTMLButtonElement {
		const control = doc.createElement('button');
		control.type = 'button';
		control.textContent = caption;
		control.addEventListener('click', action, { signal });
		return control;
	}
	const handle = button(text('pptx.insert.recordingControls', 'Recording controls'), () =>
		setExpanded(panel.hidden === true),
	);
	handle.className = 'pptx-recording-handle';
	const indicator = doc.createElement('span');
	indicator.className = 'pptx-recording-indicator';
	indicator.setAttribute('aria-hidden', 'true');
	handle.prepend(indicator);
	const elapsed = doc.createElement('span');
	elapsed.className = 'pptx-recording-elapsed';
	elapsed.setAttribute('role', 'timer');
	elapsed.setAttribute('aria-label', text('pptx.presenter.elapsed', 'Elapsed'));
	const strip = doc.createElement('div');
	strip.className = 'pptx-recording-strip';
	strip.append(handle, elapsed);
	const pauseLabel = text('pptx.rehearse.pause', 'Pause');
	const resumeLabel = text('pptx.rehearse.resume', 'Resume');
	const pause = button(pauseLabel, () => {
		if (recorder.state === 'recording') recorder.pause();
		else if (recorder.state === 'paused') recorder.resume();
	});
	const stop = button(text('pptx.followMode.stop', 'Stop'), () => {
		if (recorder.state !== 'inactive') recorder.stop();
	});
	const cancel = button(text('pptx.common.cancel', 'Cancel'), discard);
	let pinned = false;
	const pin = button(text('pptx.insert.pinRecordingDock', 'Pin dock'), () => {
		pinned = !pinned;
		pin.setAttribute('aria-pressed', String(pinned));
		scheduleCollapse();
	});
	pin.setAttribute('aria-pressed', 'false');
	const panel = doc.createElement('div');
	panel.className = 'pptx-recording-actions';
	panel.append(pause, stop, pin, cancel);
	dock.append(style, strip, panel);

	let collapseTimer: number | undefined;
	function setExpanded(expanded: boolean) {
		if (!expanded && panel.matches(':focus-within')) handle.focus();
		panel.hidden = !expanded;
		handle.setAttribute('aria-expanded', String(expanded));
	}
	function scheduleCollapse() {
		view?.clearTimeout(collapseTimer);
		if (
			!signal.aborted &&
			canHover &&
			!pinned &&
			!dock.matches(':hover') &&
			!panel.matches(':focus-within')
		) {
			collapseTimer = view?.setTimeout(() => setExpanded(false), 2200);
		}
	}
	dock.addEventListener(
		'pointerenter',
		() => {
			if (!canHover) return;
			view?.clearTimeout(collapseTimer);
			setExpanded(true);
		},
		{ signal },
	);
	dock.addEventListener('pointerleave', scheduleCollapse, { signal });
	dock.addEventListener('focusin', scheduleCollapse, { signal });
	dock.addEventListener('focusout', () => queueMicrotask(scheduleCollapse), { signal });

	let runningSince: number | undefined = performance.now();
	let recordedMs = 0;
	function syncTime() {
		const ongoingMs = runningSince === undefined ? 0 : performance.now() - runningSince;
		elapsed.textContent = formatMediaTime(Math.floor((recordedMs + ongoingMs) / 1000)).slice(0, -2);
	}
	recorder.addEventListener(
		'pause',
		() => {
			if (runningSince !== undefined) recordedMs += performance.now() - runningSince;
			runningSince = undefined;
			dock.dataset.paused = '';
			pause.textContent = resumeLabel;
			syncTime();
		},
		{ signal },
	);
	recorder.addEventListener(
		'resume',
		() => {
			runningSince = performance.now();
			delete dock.dataset.paused;
			pause.textContent = pauseLabel;
		},
		{ signal },
	);
	const clock = view?.setInterval(() => {
		if (!host.isConnected) discard();
		else syncTime();
	}, 250);
	view?.addEventListener('pagehide', discard, { signal, once: true });
	const previousFocus = modalActiveElement(doc);
	host.append(dock);
	if (typeof dock.showPopover === 'function') dock.showPopover();
	setExpanded(true);
	syncTime();
	handle.focus({ preventScroll: true });
	scheduleCollapse();
	return () => {
		const focused = dock.matches(':focus-within');
		listeners.abort();
		view?.clearInterval(clock);
		view?.clearTimeout(collapseTimer);
		dock.remove();
		if (focused && previousFocus instanceof HTMLElement && previousFocus.isConnected) {
			previousFocus.focus({ preventScroll: true });
		}
	};
}
