import { inkSvg } from '../../drawingml/inkml/svg';
import type { InkObject } from '../model';

/** An ink drawing as SVG markup, its strokes stretched onto its `width` x `height` box on the grid. */
export function inkObjectSvg(ink: InkObject, width: number, height: number): string {
	return inkSvg(ink.strokes, width, height, ink.width, ink.height);
}
