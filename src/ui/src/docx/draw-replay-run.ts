// Ink Replay on Word's pages: the ink is hidden and redrawn, stroke by stroke in the order it was
// drawn, on an SVG over the scrolling canvas. Also the Draw surface's styles, added to the
// editor's shadow root when the tab first draws.
import { createInkReplay, type InkReplay } from '../ink/ink-replay';
import { inkTools } from '../ink/tools';
import type { InkTarget } from './draw-targets';
import surfaceCss from './draw-surface.css?raw';

/** Adds the Draw surface's styles to the shadow root (or document) of `node`, once. */
export function injectDrawStyles(node: Node): void {
	const host = node.getRootNode();
	const root = host instanceof ShadowRoot ? host : (node.ownerDocument ?? document).head;
	if (root.querySelector('style[data-draw-surface]')) return;
	const style = document.createElement('style');
	style.dataset.drawSurface = '';
	style.textContent = surfaceCss;
	root.append(style);
}

export interface RunningReplay {
	replay: InkReplay;
	dispose(): void;
}

/** Starts Ink Replay of `targets` over `scroller`; `ended` runs when it has played through. */
export function startInkReplay(
	scroller: HTMLElement,
	targets: readonly InkTarget[],
	ended: () => void,
	canvasShown: () => boolean,
): RunningReplay | undefined {
	const origin = scroller.getBoundingClientRect();
	const strokes = targets.flatMap((target) => {
		const rect = target.element.getBoundingClientRect();
		const kx = rect.width / (target.ink.width || 1);
		const ky = rect.height / (target.ink.height || 1);
		return target.ink.strokes.map((stroke) => ({
			points: inkTools()
				.inkPathPoints(stroke.path)
				.map((point) => ({
					x: rect.left - origin.left + scroller.scrollLeft + point.x * kx,
					y: rect.top - origin.top + scroller.scrollTop + point.y * ky,
				})),
			style: { ...stroke, width: stroke.width * kx },
		}));
	});
	const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
	svg.classList.add('dve-ink-replay');
	svg.setAttribute('aria-hidden', 'true');
	svg.setAttribute('width', String(scroller.scrollWidth));
	svg.setAttribute('height', String(scroller.scrollHeight));
	injectDrawStyles(scroller);
	scroller.classList.add('dve-inking', 'dve-ink-replaying');
	scroller.append(svg);
	const dispose = () => {
		svg.remove();
		scroller.classList.remove('dve-ink-replaying');
		if (!canvasShown()) scroller.classList.remove('dve-inking');
	};
	const replay = createInkReplay(
		svg,
		strokes,
		(time, duration, playing) => {
			if (!playing && time >= duration) setTimeout(ended, 400);
		},
		inkTools(),
	);
	if (!replay) {
		dispose();
		return undefined;
	}
	replay.toggle();
	return { replay, dispose };
}
