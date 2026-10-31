import { attribute, children } from './sheet';

/** The User row that names a shape's data graphic master. */
export const DATA_GRAPHIC_DISPLAY_ROW = 'visDGDisplayFormat';

/**
 * A shape that carries a data graphic: Visio makes it a group of its callouts. Its own cells
 * (fill, line, text, size) are still the shape's, so edits that act on the shape alone, such as
 * formatting, treat it as the shape it was.
 */
export const isDataGraphicOwner = (shape: Element): boolean =>
	children(shape, 'Section').some(
		(section) =>
			attribute(section, 'N') === 'User' &&
			children(section, 'Row').some((row) => attribute(row, 'N') === DATA_GRAPHIC_DISPLAY_ROW),
	);
