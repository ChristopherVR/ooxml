// The soft shadow of text that has no outlines to blur: faint offset copies. A module of its own
// so the WordArt painter, which loads on first use, shares it without the text effects. DOM-free.
import type { WordArtShadowGeometry } from '../../drawingml/index';
import type { PdfVectorPage } from '../../pdf/write/pdf-vector';

const COPIES = 8;

/**
 * Draws `paint(dx, dy)` as a ring of faint copies around the shadow offset, together about as
 * dark as the shadow in the middle and fading over the blur: a soft shadow for text that cannot
 * be blurred. The copies carry an empty `ActualText`, so they never extract as text.
 */
export function wordArtSoftCopies(
	page: PdfVectorPage,
	shadow: WordArtShadowGeometry,
	paint: (dx: number, dy: number) => void,
): void {
	const sigma = shadow.blur / 2;
	const offsets: [number, number][] = [[shadow.dx, shadow.dy]];
	if (sigma > 0)
		for (let index = 0; index < COPIES; index++) {
			const angle = (index / COPIES) * 2 * Math.PI;
			offsets.push([shadow.dx + Math.cos(angle) * sigma, shadow.dy + Math.sin(angle) * sigma]);
		}
	const each = 1 - Math.pow(1 - Math.min(0.999, shadow.opacity), 1 / offsets.length);
	page.op('/Span <</ActualText ()>> BDC');
	page.save();
	page.alpha(each);
	for (const [dx, dy] of offsets) paint(dx, dy);
	page.restore();
	page.endMarked();
}
