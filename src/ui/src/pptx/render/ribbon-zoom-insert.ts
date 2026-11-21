import type { PptxSlide, SummaryZoomTarget, ZoomPptxElement } from 'ooxml-core/pptx';
import type { CanvasSize } from '../types';
import { generateElementId } from './element-clipboard';
import type { RibbonZoom } from './ribbon-insert-state';

export interface ZoomChoice {
	slide: PptxSlide;
	index: number;
	sectionId?: string;
}

export function zoomChoices(
	slides: readonly PptxSlide[],
	sourceIndex: number,
	kind: RibbonZoom,
): ZoomChoice[] {
	const choices: ZoomChoice[] = [];
	const sections = new Set<string>();
	for (const [index, slide] of slides.entries()) {
		if (kind !== 'section') {
			if (kind === 'summary' || index !== sourceIndex) choices.push({ slide, index });
		} else if (slide.sectionId && !sections.has(slide.sectionId)) {
			sections.add(slide.sectionId);
			if (index !== sourceIndex) choices.push({ slide, index, sectionId: slide.sectionId });
		}
	}
	return choices;
}

export function createRibbonZoom(
	kind: RibbonZoom,
	targets: readonly (Pick<ZoomChoice, 'index' | 'sectionId'> & {
		preview: string;
		title: string;
	})[],
	canvas: CanvasSize,
	name: string,
): ZoomPptxElement | undefined {
	const first = targets[0];
	if (!first) return undefined;
	const base: ZoomPptxElement = {
		id: generateElementId(),
		type: 'zoom',
		zoomType: kind,
		targetSlideIndex: first.index,
		name,
		altText: kind === 'summary' ? name : first.title,
		x: canvas.width * 0.33,
		y: canvas.height * 0.33,
		width: canvas.width * 0.34,
		height: canvas.height * 0.34,
	};
	if (kind === 'slide') return { ...base, imageData: first.preview, returnToParent: false };
	if (kind === 'section') {
		if (!first.sectionId) return undefined;
		return { ...base, imageData: first.preview, targetSectionId: first.sectionId };
	}
	const columns = Math.ceil(Math.sqrt(targets.length));
	const rows = Math.ceil(targets.length / columns);
	const width = canvas.width * 0.82;
	const height = canvas.height * 0.72;
	const x = (canvas.width - width) / 2;
	const y = (canvas.height - height) / 2;
	const tileWidth =
		Math.min(width / columns, (height / rows) * (canvas.width / canvas.height)) * 0.92;
	const tileHeight = tileWidth * (canvas.height / canvas.width);
	const summaryTargets = targets.flatMap<SummaryZoomTarget>((target, index) =>
		target.sectionId
			? [
					{
						sectionId: target.sectionId,
						targetSlideIndex: target.index,
						title: target.title,
						imageData: target.preview,
						x: x + ((index % columns) + 0.5) * (width / columns) - tileWidth / 2,
						y: y + (Math.floor(index / columns) + 0.5) * (height / rows) - tileHeight / 2,
						width: tileWidth,
						height: tileHeight,
					},
				]
			: [],
	);
	return { ...base, x, y, width, height, summaryLayout: 'fixed', summaryTargets };
}
