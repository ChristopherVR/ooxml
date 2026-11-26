import type { SlideInkStroke } from 'ooxml-core/pptx/editor/ink';
import { inkReplayProgress, slideInkReplayTimeline } from 'ooxml-core/pptx/editor/ink';

import { drawOverlay, inkOverlayPath } from './draw-surface';

/** Replay complete owners only, so budget or path limits never hide remaining ink. */
function completeReplayStrokes(collected: readonly SlideInkStroke[]): SlideInkStroke[] {
	const counts = new Map<SlideInkStroke['owner'], number>();
	for (const { owner } of collected) counts.set(owner, (counts.get(owner) ?? 0) + 1);
	return collected.filter(({ owner }) => {
		const count =
			owner.type === 'ink'
				? owner.inkPaths.length
				: owner.type === 'contentPart'
					? owner.inkStrokes?.length
					: undefined;
		return count === counts.get(owner);
	});
}

export function createDrawReplay(
	stage: HTMLElement,
	collected: readonly SlideInkStroke[],
	onUpdate: (time: number, duration: number, playing: boolean) => void,
) {
	const win = stage.ownerDocument.defaultView;
	const strokes = completeReplayStrokes(collected);
	if (!win || !strokes.length) {
		return undefined;
	}
	const svg = drawOverlay(stage);
	svg.dataset.pptxDrawReplay = '';
	const ids = new Set(strokes.map(({ owner }) => owner.id));
	const hidden = [...stage.querySelectorAll<HTMLElement>('[data-element-id]')].flatMap((node) => {
		if (!ids.has(node.dataset.elementId ?? '')) {
			return [];
		}
		const visibility = node.style.visibility;
		node.style.visibility = 'hidden';
		return [{ node, visibility }];
	});
	const steps = slideInkReplayTimeline(strokes.map(({ style }) => style));
	const paths = strokes.map((stroke) => {
		const path = inkOverlayPath(svg, stroke.points);
		if (/^#[\da-f]{3,8}$/iu.test(stroke.style.color)) {
			path.setAttribute('stroke', stroke.style.color);
		} else {
			path.setAttribute('stroke', '#000000');
		}
		path.setAttribute('stroke-width', String(Math.min(64, Math.max(0.1, stroke.style.width))));
		path.setAttribute('stroke-opacity', String(stroke.style.opacity));
		path.setAttribute('pathLength', '1');
		path.setAttribute('stroke-dasharray', '1');
		path.setAttribute('visibility', 'hidden');
		return { path, sample: inkReplayProgress(stroke.points, stroke.style.pointTimestamps) };
	});
	const duration = steps.reduce(
		(end, step) => Math.max(end, step.startOffsetMs + step.durationMs),
		0,
	);
	let time = 0;
	let frame: number | undefined;
	let previous = 0;
	let playing = false;
	let active = 0;
	const draw = (index: number, amount: number) => {
		const step = steps[index];
		const stroke = step && paths[step.strokeIndex];
		if (!stroke) return;
		const length = stroke.sample?.(amount) ?? amount;
		stroke.path.setAttribute('visibility', length > 0 ? 'visible' : 'hidden');
		stroke.path.setAttribute('stroke-dashoffset', String(1 - length));
	};
	const paint = () => {
		let previousStep = steps[active - 1];
		while (previousStep && time < previousStep.startOffsetMs + previousStep.durationMs) {
			draw(active, 0);
			active--;
			previousStep = steps[active - 1];
		}
		let step = steps[active];
		while (step && time >= step.startOffsetMs + step.durationMs) {
			draw(active, 1);
			active++;
			step = steps[active];
		}
		if (step) draw(active, Math.max(0, (time - step.startOffsetMs) / step.durationMs));
		onUpdate(time, duration, playing);
	};
	const advance = (now: number) => {
		time = Math.min(duration, time + now - previous);
		previous = now;
		playing = time < duration;
		frame = playing ? win.requestAnimationFrame(advance) : undefined;
		paint();
	};
	const pause = () => {
		if (frame !== undefined) {
			win.cancelAnimationFrame(frame);
			frame = undefined;
		}
		playing = false;
	};
	paint();
	return {
		toggle() {
			if (playing) {
				pause();
			} else {
				if (time === duration) {
					time = 0;
				}
				playing = true;
				previous = win.performance.now();
				frame = win.requestAnimationFrame(advance);
			}
			paint();
		},
		seek(value: number) {
			if (Number.isFinite(value)) {
				time = Math.max(0, Math.min(duration, value));
				paint();
			}
		},
		dispose() {
			pause();
			svg.remove();
			for (const { node, visibility } of hidden) {
				node.style.visibility = visibility;
			}
		},
	};
}
