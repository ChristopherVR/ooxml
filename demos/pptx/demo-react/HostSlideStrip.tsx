import { SlideThumbnail } from 'pptx-react-viewer';
import type { SlidesPane } from 'pptx-react-viewer';
import type { PptxElement } from 'pptx-viewer-core';
import React from 'react';

import {
	HOST_STRIP_LABEL,
	HOST_STRIP_STYLE,
	HOST_STRIP_THUMBNAIL_WIDTH,
	hostStripItemLabel,
	hostStripItemStyle,
} from '../shared/host-thumbnail-strip';

/** Stable, so a slide without template elements keeps its memoised thumbnail. */
const NO_TEMPLATE_ELEMENTS: PptxElement[] = [];

/** A host-owned thumbnail strip, fed from the props the built-in slides pane takes. */
export function HostSlideStrip(props: React.ComponentProps<typeof SlidesPane>) {
	const { slides, templateElementsBySlideId, canvasSize, fieldContext, tableStyleContext } = props;
	return (
		<nav
			aria-label={HOST_STRIP_LABEL}
			data-host-slide-strip
			style={HOST_STRIP_STYLE as React.CSSProperties}
		>
			{slides.map((slide, index) => (
				<button
					key={slide.id}
					type='button'
					aria-label={hostStripItemLabel(index)}
					aria-current={index === props.activeSlideIndex ? 'true' : undefined}
					style={hostStripItemStyle(index === props.activeSlideIndex) as React.CSSProperties}
					onClick={() => props.onSelectSlide(index)}
				>
					<SlideThumbnail
						slide={slide}
						templateElements={templateElementsBySlideId[slide.id] ?? NO_TEMPLATE_ELEMENTS}
						canvasSize={canvasSize}
						previewWidth={HOST_STRIP_THUMBNAIL_WIDTH}
						fieldContext={fieldContext}
						tableStyleContext={tableStyleContext}
					/>
				</button>
			))}
		</nav>
	);
}
