/**
 * What the SVG-to-PDF converter (`pdf-svg.ts`) simplifies, as stable codes. A product that
 * translates its report (PowerPoint's PDF notice) maps the code to its own dictionary key; the
 * English sentence here is what `convertSvgPage` reports to products that show it as is (Visio's
 * Export page).
 */
export const PDF_SVG_NOTES = {
	gradientOpacityAveraged: 'Gradient transparency is averaged: a PDF shading has one opacity.',
	gradientSpreadApproximated: 'Repeating and reflecting gradients are drawn as one gradient.',
	gradientFlatFill: 'A gradient with no usable area is drawn as one flat colour (the middle stop).',
	gradientOutlineFlat: 'Gradient outlines are drawn as one flat colour (the middle stop).',
	shadowBlurDropped: 'Shadow blur is not applied: shadows are drawn hard-edged.',
	glowApproximated: 'Glow is approximated by a hard-edged halo at half strength (no blur).',
	reflectionDropped: 'Reflections and masked shapes are not drawn.',
	bevelDropped: 'Bevel profiles and lighting are not drawn.',
	groupPerspectiveAffine:
		'The children of a group turned with perspective use one affine fit of the projection.',
	solid3dFlat:
		'Perspective and extrusion are drawn from the projected outline with flat shading; text and pictures use one affine fit of the projection (no convergence).',
	perspectiveParallel: '3-D rotation is drawn as a parallel projection without perspective.',
	filterDropped: 'Blur and soft-edge filters are not applied.',
} as const;

/** One simplification the SVG-to-PDF converter made. */
export type PdfSvgNoteCode = keyof typeof PDF_SVG_NOTES;

/** Every code, in a stable order. */
export const PDF_SVG_NOTE_CODES = Object.keys(PDF_SVG_NOTES) as PdfSvgNoteCode[];
