/**
 * English copy of File > Export > PDF: the options dialog (Range, Include hidden slides), the
 * notice summarising what a vector PDF approximated, and one message per diagnostic code
 * (`export/pdf-export-diagnostics.ts`).
 */
export const pdfExportTranslationsEn: Record<string, string> = {
	'pptx.compatibility.partSeparator': '; ',
	'pptx.compatibility.details': 'Details',
	'pptx.pdfExport.dialog.title': 'PDF Options',
	'pptx.pdfExport.dialog.range': 'Range',
	'pptx.pdfExport.dialog.all': 'All',
	'pptx.pdfExport.dialog.current': 'Current slide',
	'pptx.pdfExport.dialog.slides': 'Slides',
	'pptx.pdfExport.dialog.from': 'From:',
	'pptx.pdfExport.dialog.to': 'To:',
	'pptx.pdfExport.dialog.includeHidden': 'Include hidden slides',
	'pptx.pdfExport.dialog.export': 'Export',
	'pptx.pdfExport.dialog.rangeInvalid':
		'Enter slide numbers from 1 to {{count}}, with the first no greater than the last.',
	'pptx.pdfExport.dialog.rangeEmpty':
		'There are no slides to export in this range. Include hidden slides or choose other slides.',
	'pptx.pdfExport.summary.message': 'PDF saved with approximations: {{parts}}.',
	'pptx.pdfExport.summary.pictures.one': '1 element was saved as a picture',
	'pptx.pdfExport.summary.pictures.other': '{{count}} elements were saved as pictures',
	'pptx.pdfExport.summary.slidesAsPictures.one': '1 slide was saved as a single picture',
	'pptx.pdfExport.summary.slidesAsPictures.other': '{{count}} slides were saved as single pictures',
	'pptx.pdfExport.summary.fonts.one': '1 font was substituted',
	'pptx.pdfExport.summary.fonts.other': '{{count}} fonts were substituted',
	'pptx.pdfExport.summary.omitted.one': '1 item could not be saved and was left out',
	'pptx.pdfExport.summary.omitted.other': '{{count}} items could not be saved and were left out',
	'pptx.pdfExport.summary.simplified.one': '1 effect was simplified',
	'pptx.pdfExport.summary.simplified.other': '{{count}} effects were simplified',
	'pptx.pdfExport.diagnostic.filterAsPicture':
		'Soft edges, blur, glow and other filter effects are drawn as pictures.',
	'pptx.pdfExport.diagnostic.blendAsPicture': 'Blend modes and masks are drawn as pictures.',
	'pptx.pdfExport.diagnostic.threeDAsPicture':
		'3-D rotation and perspective are drawn as pictures.',
	'pptx.pdfExport.diagnostic.mediaAsPicture':
		'Video, 3-D models and canvas drawings are drawn as pictures of their current frame.',
	'pptx.pdfExport.diagnostic.textShadowAsPicture':
		'Text shadows are drawn as pictures under the searchable text.',
	'pptx.pdfExport.diagnostic.svgAsPicture':
		'Drawings this export cannot write as vector graphics are drawn as pictures.',
	'pptx.pdfExport.diagnostic.fillTextAsPicture':
		'Text filled with a gradient or picture is drawn as a picture.',
	'pptx.pdfExport.diagnostic.boxFillAsPicture':
		'Fills this export cannot write as vector graphics (patterns, picture tiles, radial and semi-transparent gradients) are drawn as pictures.',
	'pptx.pdfExport.diagnostic.borderAsPicture':
		'Borders with styles other than solid, dashed or dotted, or with differing sides on a rounded box, are drawn as pictures.',
	'pptx.pdfExport.diagnostic.shadowAsPicture': 'Shadows are drawn as pictures.',
	'pptx.pdfExport.diagnostic.verticalTextAsPicture': 'Vertical text is drawn as a picture.',
	'pptx.pdfExport.diagnostic.textColourAsPicture':
		'Text in a colour this export cannot write is drawn as a picture.',
	'pptx.pdfExport.diagnostic.complexScriptAsPicture':
		'Right-to-left text and scripts that join or combine characters are drawn as pictures, with searchable text over them.',
	'pptx.pdfExport.diagnostic.slideTooMuchText':
		'A slide has more text than this export reads, so it is drawn as one picture.',
	'pptx.pdfExport.diagnostic.slideTooManyEffects':
		'A slide has too many effects to draw separately, so it is drawn as one picture.',
	'pptx.pdfExport.diagnostic.slideNotCaptured': 'A slide could not be captured and is left blank.',
	'pptx.pdfExport.diagnostic.backgroundPictureMissing':
		'A background picture could not be read and is left out.',
	'pptx.pdfExport.diagnostic.pictureMissing': 'A picture could not be read and is left out.',
	'pptx.pdfExport.diagnostic.effectMissing':
		'An effect could not be captured as a picture and is left out.',
	'pptx.pdfExport.diagnostic.drawingPartMissing':
		'Part of a drawing could not be written as vector graphics and is left out.',
	'pptx.pdfExport.diagnostic.unsearchableText':
		'Some text uses characters no available font has, so it is left out.',
	'pptx.pdfExport.diagnostic.fontStandard':
		'Text in "{{family}}" uses a standard PDF font ({{substitute}}) because the font file is not available to embed; positions match the slide, letter shapes differ.',
	'pptx.pdfExport.diagnostic.fontStandardDefaultFamily':
		'Text in the default font uses a standard PDF font ({{substitute}}) because the font file is not available to embed; positions match the slide, letter shapes differ.',
	'pptx.pdfExport.diagnostic.fontSubstituted':
		'Text in "{{family}}" uses the embedded font "{{substitute}}" instead.',
	'pptx.pdfExport.diagnostic.fontSubstitutedDefaultFamily':
		'Text in the default font uses the embedded font "{{substitute}}" instead.',
	'pptx.pdfExport.diagnostic.searchableText':
		'Text inside pictures stays searchable as invisible text.',
	'pptx.pdfExport.diagnostic.groupOpacity':
		'Transparency set on a group is applied to each of its objects separately.',
	'pptx.pdfExport.diagnostic.gradientOpacityAveraged':
		'Gradient transparency is averaged: a PDF shading has one opacity.',
	'pptx.pdfExport.diagnostic.gradientSpreadApproximated':
		'Repeating and reflecting gradients are drawn as one gradient.',
	'pptx.pdfExport.diagnostic.gradientFlatFill':
		'A gradient with no usable area is drawn as one flat colour (the middle stop).',
	'pptx.pdfExport.diagnostic.gradientOutlineFlat':
		'Gradient outlines are drawn as one flat colour (the middle stop).',
	'pptx.pdfExport.diagnostic.shadowBlurDropped':
		'Shadow blur is not applied: shadows are drawn hard-edged.',
	'pptx.pdfExport.diagnostic.glowApproximated':
		'Glow is approximated by a hard-edged halo at half strength (no blur).',
	'pptx.pdfExport.diagnostic.reflectionDropped': 'Reflections and masked shapes are not drawn.',
	'pptx.pdfExport.diagnostic.bevelDropped': 'Bevel profiles and lighting are not drawn.',
	'pptx.pdfExport.diagnostic.groupPerspectiveAffine':
		'The children of a group turned with perspective use one affine fit of the projection.',
	'pptx.pdfExport.diagnostic.solid3dFlat':
		'Perspective and extrusion are drawn from the projected outline with flat shading; text and pictures use one affine fit of the projection.',
	'pptx.pdfExport.diagnostic.perspectiveParallel':
		'3-D rotation is drawn as a parallel projection without perspective.',
	'pptx.pdfExport.diagnostic.filterDropped': 'Blur and soft-edge filters are not applied.',
};
