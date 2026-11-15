/** Archivo > Exportar > PDF: cuadro de opciones, aviso de aproximaciones y mensajes de diagnóstico. */
export const translations: Record<string, string> = {
	'pptx.compatibility.partSeparator': '; ',
	'pptx.compatibility.details': 'Detalles',
	'pptx.pdfExport.dialog.title': 'Opciones de PDF',
	'pptx.pdfExport.dialog.range': 'Intervalo',
	'pptx.pdfExport.dialog.all': 'Todas',
	'pptx.pdfExport.dialog.current': 'Diapositiva actual',
	'pptx.pdfExport.dialog.slides': 'Diapositivas',
	'pptx.pdfExport.dialog.from': 'Desde:',
	'pptx.pdfExport.dialog.to': 'Hasta:',
	'pptx.pdfExport.dialog.includeHidden': 'Incluir diapositivas ocultas',
	'pptx.pdfExport.dialog.export': 'Exportar',
	'pptx.pdfExport.dialog.rangeInvalid':
		'Escriba números de diapositiva del 1 al {{count}}; el primero no puede ser mayor que el último.',
	'pptx.pdfExport.dialog.rangeEmpty':
		'No hay diapositivas que exportar en este intervalo. Incluya las diapositivas ocultas o elija otras diapositivas.',
	'pptx.pdfExport.summary.message': 'PDF guardado con aproximaciones: {{parts}}.',
	'pptx.pdfExport.summary.pictures.one': '1 elemento se guardó como imagen',
	'pptx.pdfExport.summary.pictures.other': '{{count}} elementos se guardaron como imágenes',
	'pptx.pdfExport.summary.slidesAsPictures.one': '1 diapositiva se guardó como una sola imagen',
	'pptx.pdfExport.summary.slidesAsPictures.other':
		'{{count}} diapositivas se guardaron como imágenes únicas',
	'pptx.pdfExport.summary.fonts.one': '1 fuente se sustituyó',
	'pptx.pdfExport.summary.fonts.other': '{{count}} fuentes se sustituyeron',
	'pptx.pdfExport.summary.omitted.one': 'se omitió contenido en 1 diapositiva',
	'pptx.pdfExport.summary.omitted.other': 'se omitió contenido en {{count}} diapositivas',
	'pptx.pdfExport.summary.simplified.one': 'se simplificaron efectos en 1 diapositiva',
	'pptx.pdfExport.summary.simplified.other': 'se simplificaron efectos en {{count}} diapositivas',
	'pptx.pdfExport.diagnostic.filterAsPicture':
		'Los bordes suaves, el desenfoque, el resplandor y otros efectos de filtro se dibujan como imágenes.',
	'pptx.pdfExport.diagnostic.blendAsPicture':
		'Los modos de fusión y las máscaras se dibujan como imágenes.',
	'pptx.pdfExport.diagnostic.threeDAsPicture':
		'El giro 3D y la perspectiva se dibujan como imágenes.',
	'pptx.pdfExport.diagnostic.mediaAsPicture':
		'Los vídeos, los modelos 3D y los dibujos de canvas se dibujan como imágenes de su fotograma actual.',
	'pptx.pdfExport.diagnostic.textShadowAsPicture':
		'Las sombras del texto se dibujan como imágenes bajo el texto que se puede buscar.',
	'pptx.pdfExport.diagnostic.svgAsPicture':
		'Los dibujos que esta exportación no puede escribir como gráficos vectoriales se dibujan como imágenes.',
	'pptx.pdfExport.diagnostic.fillTextAsPicture':
		'El texto con relleno de degradado o de imagen se dibuja como imagen.',
	'pptx.pdfExport.diagnostic.boxFillAsPicture':
		'Los rellenos que esta exportación no puede escribir como gráficos vectoriales (tramas, mosaicos de imagen, degradados radiales y semitransparentes) se dibujan como imágenes.',
	'pptx.pdfExport.diagnostic.borderAsPicture':
		'Los bordes con estilos distintos de sólido, discontinuo o punteado, o con lados distintos en un cuadro redondeado, se dibujan como imágenes.',
	'pptx.pdfExport.diagnostic.shadowAsPicture': 'Las sombras se dibujan como imágenes.',
	'pptx.pdfExport.diagnostic.verticalTextAsPicture': 'El texto vertical se dibuja como imagen.',
	'pptx.pdfExport.diagnostic.textColourAsPicture':
		'El texto en un color que esta exportación no puede escribir se dibuja como imagen.',
	'pptx.pdfExport.diagnostic.complexScriptAsPicture':
		'El texto de derecha a izquierda y los alfabetos que unen o combinan caracteres se dibujan como imágenes, con el texto que se puede buscar encima.',
	'pptx.pdfExport.diagnostic.noFontTextAsPicture':
		'El texto con caracteres que ninguna fuente del archivo puede mostrar (como chino, japonés o coreano cuando la presentación no incrusta una fuente para ellos) se dibuja como una imagen, con el texto que se puede buscar encima.',
	'pptx.pdfExport.diagnostic.slideTooMuchText':
		'Una diapositiva tiene más texto del que lee esta exportación, así que se dibuja como una sola imagen.',
	'pptx.pdfExport.diagnostic.slideTooManyEffects':
		'Una diapositiva tiene demasiados efectos para dibujarlos por separado, así que se dibuja como una sola imagen.',
	'pptx.pdfExport.diagnostic.slideNotCaptured':
		'No se pudo capturar una diapositiva y queda en blanco.',
	'pptx.pdfExport.diagnostic.backgroundPictureMissing':
		'No se pudo leer una imagen de fondo y se omite.',
	'pptx.pdfExport.diagnostic.pictureMissing': 'No se pudo leer una imagen y se omite.',
	'pptx.pdfExport.diagnostic.effectMissing':
		'No se pudo capturar un efecto como imagen y se omite.',
	'pptx.pdfExport.diagnostic.drawingPartMissing':
		'No se pudo escribir parte de un dibujo como gráfico vectorial y se omite.',
	'pptx.pdfExport.diagnostic.unsearchableText':
		'Parte del texto usa caracteres que no tiene ninguna fuente disponible, así que se omite.',
	'pptx.pdfExport.diagnostic.fontStandard':
		'El texto en «{{family}}» usa una fuente PDF estándar ({{substitute}}) porque el archivo de fuente no está disponible para incrustarlo; las posiciones coinciden con la diapositiva, pero la forma de las letras es distinta.',
	'pptx.pdfExport.diagnostic.fontStandardDefaultFamily':
		'El texto en la fuente predeterminada usa una fuente PDF estándar ({{substitute}}) porque el archivo de fuente no está disponible para incrustarlo; las posiciones coinciden con la diapositiva, pero la forma de las letras es distinta.',
	'pptx.pdfExport.diagnostic.fontSubstituted':
		'El texto en «{{family}}» usa en su lugar la fuente incrustada «{{substitute}}».',
	'pptx.pdfExport.diagnostic.fontSubstitutedDefaultFamily':
		'El texto en la fuente predeterminada usa en su lugar la fuente incrustada «{{substitute}}».',
	'pptx.pdfExport.diagnostic.searchableText':
		'El texto dentro de las imágenes se puede seguir buscando como texto invisible.',
	'pptx.pdfExport.diagnostic.groupOpacity':
		'La transparencia de un grupo se aplica a cada uno de sus objetos por separado.',
	'pptx.pdfExport.diagnostic.gradientOpacityAveraged':
		'La transparencia de los degradados se promedia: un sombreado PDF tiene una sola opacidad.',
	'pptx.pdfExport.diagnostic.gradientSpreadApproximated':
		'Los degradados repetidos y reflejados se dibujan como un solo degradado.',
	'pptx.pdfExport.diagnostic.gradientFlatFill':
		'Un degradado sin área utilizable se dibuja como un solo color plano (la parada central).',
	'pptx.pdfExport.diagnostic.gradientOutlineFlat':
		'Los contornos con degradado se dibujan como un solo color plano (la parada central).',
	'pptx.pdfExport.diagnostic.shadowBlurDropped':
		'No se aplica el desenfoque de las sombras: las sombras se dibujan con bordes nítidos.',
	'pptx.pdfExport.diagnostic.glowApproximated':
		'El resplandor se aproxima con un halo de bordes nítidos a media intensidad (sin desenfoque).',
	'pptx.pdfExport.diagnostic.reflectionDropped':
		'No se dibujan los reflejos ni las formas con máscara.',
	'pptx.pdfExport.diagnostic.bevelDropped':
		'No se dibujan los perfiles de bisel ni la iluminación.',
	'pptx.pdfExport.diagnostic.groupPerspectiveAffine':
		'Los elementos de un grupo girado con perspectiva usan un ajuste afín de la proyección.',
	'pptx.pdfExport.diagnostic.solid3dFlat':
		'La perspectiva y la extrusión se dibujan a partir del contorno proyectado con sombreado plano; el texto y las imágenes usan un ajuste afín de la proyección.',
	'pptx.pdfExport.diagnostic.perspectiveParallel':
		'El giro 3D se dibuja como una proyección paralela sin perspectiva.',
	'pptx.pdfExport.diagnostic.filterDropped':
		'No se aplican los filtros de desenfoque ni de bordes suaves.',
};
