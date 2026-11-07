import { buildPreviewElements, deviceStrokeStageStyle, slideThumbnailLayout } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import React from 'react';

import { SLIDE_NAV_THUMBNAIL_WIDTH, SLIDE_TRANSITION_OPTIONS } from '../constants';
import { useTranslation } from '../i18n/use-translation';
import type { CanvasSize } from '../types';
import { getReactSlideBackgroundStyle } from '../utils/slide-background-style';
import type { TableStyleContext } from '../utils/table-band-style';
import type { FieldSubstitutionContext } from '../utils/text-field-substitution';
import { deriveSlideFieldContext } from './slide-field-context';
import { SlideBackgroundImageLayer } from './SlideBackgroundImageLayer';
import { StaticElementRenderer } from './StaticElementRenderer';

/**
 * A static, scaled-down drawing of one slide, as the slides pane shows it. It renders from the
 * slide model, so it follows unsaved edits. Feed it from `slidesPaneProps`:
 * `templateElementsBySlideId[slide.id]`, `canvasSize`, `fieldContext` and `tableStyleContext`.
 */
export interface SlideThumbnailProps {
	slide: PptxSlide;
	/** The slide's inherited layout and master elements, drawn beneath its own. */
	templateElements: PptxElement[];
	canvasSize: CanvasSize;
	/**
	 * Thumbnail width in pixels; the height follows the deck's aspect ratio. Set it to fill a host
	 * strip without the pane's frame. Omitted, the thumbnail is the pane width with a border.
	 */
	previewWidth?: number;
	/** Presentation-wide field context (date/header/footer/custom props). */
	fieldContext?: FieldSubstitutionContext;
	/** Theme + table style map for resolving table band/header colours. */
	tableStyleContext?: TableStyleContext;
}

function SlideThumbnailImpl({
	slide,
	templateElements,
	canvasSize,
	previewWidth,
	fieldContext,
	tableStyleContext,
}: SlideThumbnailProps): React.ReactElement {
	const safeCanvasWidth = Math.max(canvasSize.width, 1);
	const safeCanvasHeight = Math.max(canvasSize.height, 1);
	const { height: previewHeight, scale } = slideThumbnailLayout(
		safeCanvasWidth,
		safeCanvasHeight,
		previewWidth ?? SLIDE_NAV_THUMBNAIL_WIDTH,
	);
	const previewElements = buildPreviewElements(slide, templateElements);
	const slideFieldContext = deriveSlideFieldContext(fieldContext, slide);
	const { t } = useTranslation();
	const transitionLabel = slide.transition
		? (() => {
				const opt = SLIDE_TRANSITION_OPTIONS.find((o) => o.value === slide.transition?.type);
				return opt ? t(opt.i18nKey) : slide.transition.type;
			})()
		: undefined;

	return (
		<div
			className={
				previewWidth === undefined
					? 'relative w-full overflow-hidden rounded border border-border bg-white'
					: 'relative w-full overflow-hidden bg-white'
			}
			style={{
				height: previewHeight,
				...getReactSlideBackgroundStyle(slide, {
					widthPx: safeCanvasWidth,
					heightPx: safeCanvasHeight,
				}),
			}}
		>
			<div
				className='absolute top-0 left-0 origin-top-left'
				style={{
					width: safeCanvasWidth,
					height: safeCanvasHeight,
					transform: `scale(${scale})`,
					transformOrigin: 'top left',
					...deviceStrokeStageStyle(scale),
				}}
			>
				<SlideBackgroundImageLayer slide={slide} />
				{/* Transition indicator badge */}
				{slide.transition &&
					slide.transition.type !== 'none' &&
					slide.transition.type !== 'cut' && (
						<div
							className='absolute top-0.5 right-0.5 z-10 px-1 py-px rounded bg-primary/80 text-[7px] text-white leading-tight pointer-events-none'
							title={t('pptx.slideThumbnail.transitionTooltip', { name: transitionLabel })}
						>
							{transitionLabel}
						</div>
					)}
				{previewElements.map((element, index) => (
					<StaticElementRenderer
						key={element.id}
						element={element}
						activeSlide={slide}
						allSlides={[slide]}
						zIndex={index}
						fieldContext={slideFieldContext}
						tableStyleContext={tableStyleContext}
					/>
				))}
			</div>
		</div>
	);
}

/**
 * Memo comparator: re-render only when the slide identity, dirty/hidden flags,
 * underlying elements reference, template elements reference, or canvas size
 * changes. Avoids burning frames on parent re-renders that don't change props.
 */
export function slideThumbnailPropsEqual(
	prev: SlideThumbnailProps,
	next: SlideThumbnailProps,
): boolean {
	if (prev.previewWidth !== next.previewWidth) {
		return false;
	}
	if (prev.slide.id !== next.slide.id) {
		return false;
	}
	if (prev.slide.isDirty !== next.slide.isDirty) {
		return false;
	}
	if (prev.slide.hidden !== next.slide.hidden) {
		return false;
	}
	if (prev.slide.elements !== next.slide.elements) {
		return false;
	}
	// `buildPreviewElements` -> `visibleTemplateElements` drops the inherited
	// layout/master artwork entirely when this is `false`, so toggling "Hide
	// Background Graphics" must invalidate the memo even though `elements`
	// and `templateElements` themselves haven't changed.
	if (prev.slide.showMasterShapes !== next.slide.showMasterShapes) {
		return false;
	}
	if (prev.slide.backgroundColor !== next.slide.backgroundColor) {
		return false;
	}
	if (prev.slide.backgroundImage !== next.slide.backgroundImage) {
		return false;
	}
	if (prev.slide.backgroundGradient !== next.slide.backgroundGradient) {
		return false;
	}
	if (prev.slide.backgroundPattern !== next.slide.backgroundPattern) {
		return false;
	}
	if (prev.templateElements !== next.templateElements) {
		return false;
	}
	if (prev.fieldContext !== next.fieldContext) {
		return false;
	}
	if (prev.tableStyleContext !== next.tableStyleContext) {
		return false;
	}
	if (
		prev.canvasSize.width !== next.canvasSize.width ||
		prev.canvasSize.height !== next.canvasSize.height
	) {
		return false;
	}
	return true;
}

export const SlideThumbnail = /* @__PURE__ */ React.memo(
	SlideThumbnailImpl,
	slideThumbnailPropsEqual,
);
