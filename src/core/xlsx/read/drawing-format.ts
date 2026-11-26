import { parseDrawingEffects } from '../../drawingml/drawing-effects';
import type { DrawingOuterShadow } from '../../drawingml/effect-model';
import { NS, first, type XmlElement } from '../../xml/index';

export function readDrawingShadow(props: XmlElement | undefined): DrawingOuterShadow | undefined {
	return parseDrawingEffects(first(props, 'effectLst', NS.a))?.effects.find(
		(effect) => effect.kind === 'outerShadow',
	);
}
