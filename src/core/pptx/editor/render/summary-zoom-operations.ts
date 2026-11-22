import { remapZoomTargets } from '../../core/utils/remap-zoom-targets';
import type { PptxSection, PptxSlide } from '../../core/types';
import { generateSectionId } from './section-operations';
import { createBlankSlide } from './slide-operations';

export interface SummaryZoomTargetSlide {
	slideId: string;
	title: string;
	preview: string;
}

export interface SummaryZoomMutation {
	slides: PptxSlide[];
	sections: PptxSection[];
	slideIndex: number;
}

/** Begin a section at each picked slide and insert the overview before the first pick. */
export function prepareSummaryZoom(
	slides: readonly PptxSlide[],
	sections: readonly PptxSection[],
	targets: readonly SummaryZoomTargetSlide[],
	options: { name: string; defaultSectionName: string; keepUnused: boolean },
) {
	const selected = new Map(targets.map((target) => [target.slideId, target]));
	const first = slides.findIndex((slide) => selected.has(slide.id));
	if (first < 0 || targets.some(({ slideId }) => !slides.some(({ id }) => id === slideId))) {
		return undefined;
	}

	const originalSections = new Map(sections.map((section) => [section.id, section]));
	const overviewSection: PptxSection = {
		id: generateSectionId(),
		name: options.name,
		slideIds: [],
	};
	const overview = { ...createBlankSlide(first + 1), name: options.name, showMasterShapes: false };
	const result: SummaryZoomMutation = { slides: [], sections: [], slideIndex: first };
	const entries: Array<SummaryZoomTargetSlide & { index: number; sectionId: string }> = [];
	let previousOriginal: string | undefined;
	let section: PptxSection | undefined;
	const append = (slide: PptxSlide, owner: PptxSection) => {
		owner.slideIds.push(slide.slideId ?? slide.id);
		result.slides.push({
			...slide,
			sectionId: owner.id,
			sectionName: owner.name,
			slideNumber: result.slides.length + 1,
		});
	};
	for (const [index, slide] of slides.entries()) {
		const target = selected.get(slide.id);
		const boundary = index === 0 || previousOriginal !== slide.sectionId;
		const existing = slide.sectionId ? originalSections.get(slide.sectionId) : undefined;
		previousOriginal = slide.sectionId;
		if (index === first) {
			result.sections.push(overviewSection);
			append(overview, overviewSection);
		}
		if (target || !section || (boundary && existing && options.keepUnused)) {
			const reused = boundary ? existing : undefined;
			section = reused
				? { ...reused, collapsed: target ? false : reused.collapsed, slideIds: [] }
				: {
						id: generateSectionId(),
						name: target?.title ?? options.defaultSectionName,
						slideIds: [],
					};
			result.sections.push(section);
		}
		if (target) entries.push({ ...target, index: result.slides.length, sectionId: section.id });
		append(slide, section);
	}

	// Keep empty sections when requested. They have no slide boundary of their own.
	if (options.keepUnused) {
		result.sections.push(...sections.filter(({ slideIds }) => !slideIds.length));
	}
	result.slides = remapZoomTargets(slides, result.slides);
	return { result, entries };
}
