// Ink Replay over Excel's cells: the sheet's ink is hidden and redrawn, stroke by stroke in the
// order it was drawn, on an SVG over the grid.
import { createInkReplay, type InkReplay } from '../../ink/ink-replay';
import { inkTools } from '../../ink/tools';
import type { GridView } from './grid-view';
import { sheetInks, zoomOf } from './ink-sheet';

/** Starts Ink Replay of the shown sheet's ink; `ended` runs when it has played through. */
export function startInkReplay(view: GridView, ended: () => void): InkReplay | undefined {
	const g = view.geometry;
	const f = zoomOf(view);
	const main = g.box('main');
	const { ox, oy } = g.origin('main');
	const strokes = sheetInks(view, false).flatMap(({ ink, box }) => {
		const kx = box.width / (ink.width || 1);
		const ky = box.height / (ink.height || 1);
		return ink.strokes.map((stroke) => ({
			points: inkTools()
				.inkPathPoints(stroke.path)
				.map((point) => ({
					x: main.x + (box.x + point.x * kx) * f - ox,
					y: main.y + (box.y + point.y * ky) * f - oy,
				})),
			style: { ...stroke, width: stroke.width * f },
		}));
	});
	const svg = view.doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
	svg.classList.add('xg-ink-replay');
	svg.setAttribute('aria-hidden', 'true');
	svg.setAttribute('width', String(view.view.clientWidth));
	svg.setAttribute('height', String(view.view.clientHeight));
	view.view.append(svg);
	view.root.classList.add('xg-ink-replaying');
	const replay = createInkReplay(
		svg,
		strokes,
		(time, duration, playing) => {
			if (!playing && time >= duration) setTimeout(ended, 400);
		},
		inkTools(),
	);
	if (!replay) {
		svg.remove();
		view.root.classList.remove('xg-ink-replaying');
		return undefined;
	}
	replay.toggle();
	return replay;
}
