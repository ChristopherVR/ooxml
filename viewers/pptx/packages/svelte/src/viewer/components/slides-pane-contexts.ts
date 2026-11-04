import type { CanvasSize, FieldSubstitutionContext } from 'ooxml-ui/pptx';

import { getFieldContextGetter, provideFieldContext } from '../state/field-context';
import { getRenderContextSource, provideRenderContext } from '../state/render-context';
import type { SlidesPaneTableStyleContext } from './slides-pane-props';

/** The deck props the slide thumbnails resolve fields and table styles from. */
export interface SlidesPaneContextSources {
	canvasSize: CanvasSize;
	fieldContext?: FieldSubstitutionContext | undefined;
	tableStyleContext?: SlidesPaneTableStyleContext | undefined;
}

/**
 * Provide the thumbnails' field and table style contexts from the host's props, falling back to
 * the enclosing `createViewerState` context for an omitted prop. Call during component init:
 * Svelte context is captured once, so `sources` is a getter and every provided value reads it live.
 * `SlidesPane` and `SlideThumbnail` share it, so both draw a slide the same way.
 */
export function provideSlidesPaneContexts(sources: () => SlidesPaneContextSources): void {
	const inheritedFieldContext = getFieldContextGetter();
	provideFieldContext(() => sources().fieldContext ?? inheritedFieldContext?.());
	const inherited = getRenderContextSource();
	const table = (): SlidesPaneTableStyleContext | undefined => sources().tableStyleContext;
	provideRenderContext({
		getColorScheme: () => {
			const own = table();
			return own ? own.theme?.colorScheme : inherited?.getColorScheme();
		},
		getTableStyleMap: () => {
			const own = table();
			return own ? own.tableStyleMap : inherited?.getTableStyleMap();
		},
		getFontScheme: () => {
			const own = table();
			return own ? own.theme?.fontScheme : inherited?.getFontScheme?.();
		},
		getFormatScheme: () => {
			const own = table();
			return own ? own.theme?.formatScheme : inherited?.getFormatScheme?.();
		},
		getCanvasSize: () => sources().canvasSize,
	});
}
