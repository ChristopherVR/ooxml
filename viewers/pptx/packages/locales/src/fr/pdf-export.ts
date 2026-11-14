/** Fichier > Exporter > PDF : boîte d'options, avis d'approximations et messages de diagnostic. */
export const translations: Record<string, string> = {
	'pptx.compatibility.partSeparator': ' ; ',
	'pptx.compatibility.details': 'Détails',
	'pptx.pdfExport.dialog.title': 'Options PDF',
	'pptx.pdfExport.dialog.range': 'Étendue',
	'pptx.pdfExport.dialog.all': 'Toutes',
	'pptx.pdfExport.dialog.current': 'Diapositive en cours',
	'pptx.pdfExport.dialog.slides': 'Diapositives',
	'pptx.pdfExport.dialog.from': 'De :',
	'pptx.pdfExport.dialog.to': 'À :',
	'pptx.pdfExport.dialog.includeHidden': 'Inclure les diapositives masquées',
	'pptx.pdfExport.dialog.export': 'Exporter',
	'pptx.pdfExport.dialog.rangeInvalid':
		'Entrez des numéros de diapositive de 1 à {{count}}, le premier ne dépassant pas le dernier.',
	'pptx.pdfExport.dialog.rangeEmpty':
		'Cette étendue ne contient aucune diapositive à exporter. Incluez les diapositives masquées ou choisissez d’autres diapositives.',
	'pptx.pdfExport.summary.message': 'PDF enregistré avec des approximations : {{parts}}.',
	'pptx.pdfExport.summary.pictures.one': '1 élément a été enregistré comme image',
	'pptx.pdfExport.summary.pictures.other': '{{count}} éléments ont été enregistrés comme images',
	'pptx.pdfExport.summary.slidesAsPictures.one':
		'1 diapositive a été enregistrée comme une seule image',
	'pptx.pdfExport.summary.slidesAsPictures.other':
		'{{count}} diapositives ont été enregistrées chacune comme une seule image',
	'pptx.pdfExport.summary.fonts.one': '1 police a été remplacée',
	'pptx.pdfExport.summary.fonts.other': '{{count}} polices ont été remplacées',
	'pptx.pdfExport.summary.omitted.one': '1 élément n’a pas pu être enregistré et a été omis',
	'pptx.pdfExport.summary.omitted.other':
		'{{count}} éléments n’ont pas pu être enregistrés et ont été omis',
	'pptx.pdfExport.summary.simplified.one': '1 effet a été simplifié',
	'pptx.pdfExport.summary.simplified.other': '{{count}} effets ont été simplifiés',
	'pptx.pdfExport.diagnostic.filterAsPicture':
		'Les bords adoucis, le flou, la lumière et les autres effets de filtre sont dessinés comme des images.',
	'pptx.pdfExport.diagnostic.blendAsPicture':
		'Les modes de fusion et les masques sont dessinés comme des images.',
	'pptx.pdfExport.diagnostic.threeDAsPicture':
		'La rotation 3D et la perspective sont dessinées comme des images.',
	'pptx.pdfExport.diagnostic.mediaAsPicture':
		'Les vidéos, les modèles 3D et les dessins canvas sont dessinés comme des images de leur image actuelle.',
	'pptx.pdfExport.diagnostic.textShadowAsPicture':
		'Les ombres du texte sont dessinées comme des images sous le texte interrogeable.',
	'pptx.pdfExport.diagnostic.svgAsPicture':
		'Les dessins que cet export ne peut pas écrire en graphiques vectoriels sont dessinés comme des images.',
	'pptx.pdfExport.diagnostic.fillTextAsPicture':
		'Le texte rempli d’un dégradé ou d’une image est dessiné comme une image.',
	'pptx.pdfExport.diagnostic.boxFillAsPicture':
		'Les remplissages que cet export ne peut pas écrire en graphiques vectoriels (motifs, mosaïques d’images, dégradés radiaux et semi-transparents) sont dessinés comme des images.',
	'pptx.pdfExport.diagnostic.borderAsPicture':
		'Les bordures d’un style autre que plein, tirets ou pointillés, ou aux côtés différents sur une zone arrondie, sont dessinées comme des images.',
	'pptx.pdfExport.diagnostic.shadowAsPicture': 'Les ombres sont dessinées comme des images.',
	'pptx.pdfExport.diagnostic.verticalTextAsPicture':
		'Le texte vertical est dessiné comme une image.',
	'pptx.pdfExport.diagnostic.textColourAsPicture':
		'Le texte dans une couleur que cet export ne peut pas écrire est dessiné comme une image.',
	'pptx.pdfExport.diagnostic.complexScriptAsPicture':
		'Le texte de droite à gauche et les écritures qui lient ou combinent des caractères sont dessinés comme des images, avec le texte interrogeable par-dessus.',
	'pptx.pdfExport.diagnostic.slideTooMuchText':
		'Une diapositive contient plus de texte que cet export n’en lit ; elle est donc dessinée comme une seule image.',
	'pptx.pdfExport.diagnostic.slideTooManyEffects':
		'Une diapositive a trop d’effets pour les dessiner séparément ; elle est donc dessinée comme une seule image.',
	'pptx.pdfExport.diagnostic.slideNotCaptured':
		'Une diapositive n’a pas pu être capturée et reste vide.',
	'pptx.pdfExport.diagnostic.backgroundPictureMissing':
		'Une image d’arrière-plan n’a pas pu être lue et est omise.',
	'pptx.pdfExport.diagnostic.pictureMissing': 'Une image n’a pas pu être lue et est omise.',
	'pptx.pdfExport.diagnostic.effectMissing':
		'Un effet n’a pas pu être capturé comme image et est omis.',
	'pptx.pdfExport.diagnostic.drawingPartMissing':
		'Une partie d’un dessin n’a pas pu être écrite en graphique vectoriel et est omise.',
	'pptx.pdfExport.diagnostic.unsearchableText':
		'Une partie du texte utilise des caractères qu’aucune police disponible ne contient ; elle est donc omise.',
	'pptx.pdfExport.diagnostic.fontStandard':
		'Le texte en « {{family}} » utilise une police PDF standard ({{substitute}}) car le fichier de police n’est pas disponible pour l’incorporer ; les positions correspondent à la diapositive, la forme des lettres diffère.',
	'pptx.pdfExport.diagnostic.fontStandardDefaultFamily':
		'Le texte dans la police par défaut utilise une police PDF standard ({{substitute}}) car le fichier de police n’est pas disponible pour l’incorporer ; les positions correspondent à la diapositive, la forme des lettres diffère.',
	'pptx.pdfExport.diagnostic.fontSubstituted':
		'Le texte en « {{family}} » utilise à la place la police incorporée « {{substitute}} ».',
	'pptx.pdfExport.diagnostic.fontSubstitutedDefaultFamily':
		'Le texte dans la police par défaut utilise à la place la police incorporée « {{substitute}} ».',
	'pptx.pdfExport.diagnostic.searchableText':
		'Le texte à l’intérieur des images reste interrogeable sous forme de texte invisible.',
	'pptx.pdfExport.diagnostic.groupOpacity':
		'La transparence d’un groupe est appliquée séparément à chacun de ses objets.',
	'pptx.pdfExport.diagnostic.gradientOpacityAveraged':
		'La transparence des dégradés est moyennée : un ombrage PDF n’a qu’une opacité.',
	'pptx.pdfExport.diagnostic.gradientSpreadApproximated':
		'Les dégradés répétés et réfléchis sont dessinés comme un seul dégradé.',
	'pptx.pdfExport.diagnostic.gradientFlatFill':
		'Un dégradé sans surface utilisable est dessiné d’une seule couleur unie (le point central).',
	'pptx.pdfExport.diagnostic.gradientOutlineFlat':
		'Les contours en dégradé sont dessinés d’une seule couleur unie (le point central).',
	'pptx.pdfExport.diagnostic.shadowBlurDropped':
		'Le flou des ombres n’est pas appliqué : les ombres sont dessinées avec des bords nets.',
	'pptx.pdfExport.diagnostic.glowApproximated':
		'La lumière est approchée par un halo à bords nets à mi-intensité (sans flou).',
	'pptx.pdfExport.diagnostic.reflectionDropped':
		'Les réflexions et les formes masquées ne sont pas dessinées.',
	'pptx.pdfExport.diagnostic.bevelDropped':
		'Les profils de biseau et l’éclairage ne sont pas dessinés.',
	'pptx.pdfExport.diagnostic.groupPerspectiveAffine':
		'Les éléments d’un groupe tourné en perspective utilisent une approximation affine de la projection.',
	'pptx.pdfExport.diagnostic.solid3dFlat':
		'La perspective et l’extrusion sont dessinées à partir du contour projeté avec un ombrage plat ; le texte et les images utilisent une approximation affine de la projection.',
	'pptx.pdfExport.diagnostic.perspectiveParallel':
		'La rotation 3D est dessinée en projection parallèle, sans perspective.',
	'pptx.pdfExport.diagnostic.filterDropped':
		'Les filtres de flou et de bords adoucis ne sont pas appliqués.',
};
