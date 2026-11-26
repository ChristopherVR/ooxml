import type { InkCoordinate, InkTools } from 'ooxml-core/drawingml';

/** The replay timing of the core (`loadInkTools`, or PowerPoint's own editor entry). */
export type InkReplayTiming = Pick<InkTools, 'inkReplayProgress' | 'slideInkReplayTimeline'>;

/**
 * Ink Replay, shared by PowerPoint, Word and Excel: strokes redrawn in the order and at the pace
 * they were drawn (their recorded times, else one after another) on an SVG overlay the product
 * places over its surface. The product hides the ink it replays and disposes the replay to
 * show it again.
 */
export interface InkReplayStroke {
	/** The stroke's points in the overlay's coordinates. */
	points: readonly InkCoordinate[];
	style: {
		color: string;
		width: number;
		opacity: number;
		pointTimestamps?: readonly number[] | undefined;
		startTimestamp?: number | undefined;
		durationMs?: number | undefined;
	};
}

export interface InkReplay {
	toggle(): void;
	seek(time: number): void;
	dispose(): void;
}

const SVG_NS = 'http://www.w3.org/2000/svg';

function overlayPath(svg: SVGSVGElement, points: readonly InkCoordinate[]): SVGPathElement {
	const path = svg.ownerDocument.createElementNS(SVG_NS, 'path');
	path.setAttribute('fill', 'none');
	path.setAttribute('stroke-linecap', 'round');
	path.setAttribute('stroke-linejoin', 'round');
	path.setAttribute(
		'd',
		points.map((point, index) => `${index ? 'L' : 'M'}${point.x} ${point.y}`).join(' '),
	);
	svg.append(path);
	return path;
}

/**
 * Builds the replay of `strokes` on `svg` (which `dispose` removes); `onUpdate` follows the
 * playhead for a seek bar. Undefined when there is nothing to replay or no window to animate in.
 */
export function createInkReplay(
	svg: SVGSVGElement,
	strokes: readonly InkReplayStroke[],
	onUpdate: (time: number, duration: number, playing: boolean) => void,
	{ inkReplayProgress, slideInkReplayTimeline }: InkReplayTiming,
): InkReplay | undefined {
	const win = svg.ownerDocument.defaultView;
	if (!win || !strokes.length) return undefined;
	const steps = slideInkReplayTimeline(strokes.map(({ style }) => style));
	const paths = strokes.map((stroke) => {
		const path = overlayPath(svg, stroke.points);
		path.setAttribute(
			'stroke',
			/^#[\da-f]{3,8}$/iu.test(stroke.style.color) ? stroke.style.color : '#000000',
		);
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
			if (playing) pause();
			else {
				if (time === duration) time = 0;
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
		},
	};
}
