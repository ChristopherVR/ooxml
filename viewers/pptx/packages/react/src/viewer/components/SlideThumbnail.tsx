import {
	buildPreviewElements,
	deviceStrokeStageStyle,
	slideThumbnailLayout,
	slideThumbnailTransitionBadge,
	slideThumbnailTransitionBadgeStyle,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import React from 'react';

import { SLIDE_NAV_THUMBNAIL_WIDTH } from '../constants';
import { useTranslation } from '../i18n/use-translation';
import type { CanvasSize } from '../types';
import { getReactSlideBackgroundStyle } from '../utils/slide-background-style';
import type { TableStyleContext } from '../utils/table-band-style';
import type { FieldSubstitutionContext } from '../utils/text-field-substitution';
import type { ElementRenderErrorHandler } from './ElementErrorBoundary';
import { LinkedContentScope, useLinkedContentView } from './linked-content-context';
import { deriveSlideFieldContext } from './slide-field-context';
import { SlideBackgroundImageLayer } from './SlideBackgroundImageLayer';
import { StaticElementRenderer } from './StaticElementRenderer';

/**
 * A static, scaled-down drawing of one slide, as the slides pane shows it. It renders from the
 * slide model, so it follows unsaved edits. Feed it from `slidesPaneProps`:
 * `templateElementsBySlideId[slide.id]`, `canvasSize`, `fieldContext` and `tableStyleContext`.
 */
export interface SlideThumbnailProps {
	/** Reports an element omitted from this preview; defaults to console.error. */
	onElementRenderError?: ElementRenderErrorHandler;
	/** The slide to draw. A different slide object (every edit makes one) redraws it. */
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
	/**
	 * Draw the deck's linked web content: `externalContent.allowed` from the building blocks.
	 * Unset keeps the surrounding viewer's choice, and withholds it outside a viewer.
	 */
	linkedContent?: boolean;
}

function SlideThumbnailImpl({
	slide,
	templateElements,
	canvasSize,
	previewWidth,
	fieldContext,
	tableStyleContext,
	onElementRenderError,
}: Omit<SlideThumbnailProps, 'linkedContent'>): React.ReactElement {
	const safeCanvasWidth = Math.max(canvasSize.width, 1);
	const safeCanvasHeight = Math.max(canvasSize.height, 1);
	const layout = slideThumbnailLayout(
		safeCanvasWidth,
		safeCanvasHeight,
		previewWidth ?? SLIDE_NAV_THUMBNAIL_WIDTH,
	);
	const { height: previewHeight, scale } = layout;
	const previewElements = buildPreviewElements(slide, templateElements);
	const backgroundSlide = useLinkedContentView(slide);
	const slideFieldContext = deriveSlideFieldContext(fieldContext, slide);
	const { t } = useTranslation();
	const transitionBadge = slideThumbnailTransitionBadge(slide, t);

	return (
		<div
			className={
				previewWidth === undefined
					? 'relative w-full overflow-hidden rounded border border-border bg-white'
					: 'relative w-full overflow-hidden bg-white'
			}
			style={{
				height: previewHeight,
				...getReactSlideBackgroundStyle(backgroundSlide, {
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
				{previewElements.map((element, index) => (
					<StaticElementRenderer
						key={element.id}
						element={element}
						activeSlide={slide}
						allSlides={[slide]}
						zIndex={index}
						fieldContext={slideFieldContext}
						tableStyleContext={tableStyleContext}
						onElementRenderError={onElementRenderError}
					/>
				))}
			</div>
			{transitionBadge && (
				<div
					data-pptx-thumbnail-transition=''
					title={transitionBadge.title}
					style={slideThumbnailTransitionBadgeStyle(layout)}
				>
					{transitionBadge.label}
				</div>
			)}
		</div>
	);
}

/**
 * Memo comparator: re-render only when a prop changes. The slide is compared by
 * reference: every edit replaces the slide object (the editor state is
 * immutable), and the thumbnail draws from many of its fields (elements,
 * background, transition badge, slide number and title fields, "Hide Background
 * Graphics"), so a field list would go stale the next time one is added.
 * Avoids burning frames on parent re-renders that don't change props.
 */
export function slideThumbnailPropsEqual(
	prev: SlideThumbnailProps,
	next: SlideThumbnailProps,
): boolean {
	if (prev.onElementRenderError !== next.onElementRenderError) {
		return false;
	}
	if (prev.previewWidth !== next.previewWidth) {
		return false;
	}
	if (prev.slide !== next.slide) {
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
	if (prev.linkedContent !== next.linkedContent) {
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

function SlideThumbnailScoped({ linkedContent, ...props }: SlideThumbnailProps) {
	return (
		<LinkedContentScope allowed={linkedContent}>
			<SlideThumbnailImpl {...props} />
		</LinkedContentScope>
	);
}

export const SlideThumbnail = /* @__PURE__ */ React.memo(
	SlideThumbnailScoped,
	slideThumbnailPropsEqual,
);
