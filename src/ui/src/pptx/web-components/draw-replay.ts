import type { SlideInkStroke } from 'ooxml-core/pptx/editor/ink';
import { inkReplayProgress, slideInkReplayTimeline } from 'ooxml-core/pptx/editor/ink';

import { createInkReplay } from '../../ink/ink-replay';
import { drawOverlay } from './draw-surface';

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
	const replay = createInkReplay(svg, strokes, onUpdate, {
		inkReplayProgress,
		slideInkReplayTimeline,
	});
	const restore = () => {
		for (const { node, visibility } of hidden) {
			node.style.visibility = visibility;
		}
	};
	if (!replay) {
		svg.remove();
		restore();
		return undefined;
	}
	return {
		toggle: () => replay.toggle(),
		seek: (value: number) => replay.seek(value),
		dispose() {
			replay.dispose();
			restore();
		},
	};
}
