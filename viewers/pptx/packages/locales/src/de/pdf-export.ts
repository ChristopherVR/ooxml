/** File > Exportieren > PDF: Optionsdialog, Hinweis zu Annäherungen und Diagnosemeldungen. */
export const translations: Record<string, string> = {
	'pptx.compatibility.partSeparator': '; ',
	'pptx.compatibility.details': 'Details',
	'pptx.pdfExport.dialog.title': 'PDF-Optionen',
	'pptx.pdfExport.dialog.range': 'Bereich',
	'pptx.pdfExport.dialog.all': 'Alle',
	'pptx.pdfExport.dialog.current': 'Aktuelle Folie',
	'pptx.pdfExport.dialog.slides': 'Folien',
	'pptx.pdfExport.dialog.from': 'Von:',
	'pptx.pdfExport.dialog.to': 'Bis:',
	'pptx.pdfExport.dialog.includeHidden': 'Ausgeblendete Folien einschließen',
	'pptx.pdfExport.dialog.export': 'Exportieren',
	'pptx.pdfExport.dialog.rangeInvalid':
		'Geben Sie Foliennummern von 1 bis {{count}} ein; die erste darf nicht größer als die letzte sein.',
	'pptx.pdfExport.dialog.rangeEmpty':
		'In diesem Bereich gibt es keine Folien zum Exportieren. Schließen Sie ausgeblendete Folien ein oder wählen Sie andere Folien.',
	'pptx.pdfExport.summary.message': 'PDF mit Annäherungen gespeichert: {{parts}}.',
	'pptx.pdfExport.summary.pictures.one': '1 Element wurde als Bild gespeichert',
	'pptx.pdfExport.summary.pictures.other': '{{count}} Elemente wurden als Bilder gespeichert',
	'pptx.pdfExport.summary.slidesAsPictures.one': '1 Folie wurde als einzelnes Bild gespeichert',
	'pptx.pdfExport.summary.slidesAsPictures.other':
		'{{count}} Folien wurden als einzelne Bilder gespeichert',
	'pptx.pdfExport.summary.fonts.one': '1 Schriftart wurde ersetzt',
	'pptx.pdfExport.summary.fonts.other': '{{count}} Schriftarten wurden ersetzt',
	'pptx.pdfExport.summary.omitted.one': 'auf 1 Folie wurden Inhalte weggelassen',
	'pptx.pdfExport.summary.omitted.other': 'auf {{count}} Folien wurden Inhalte weggelassen',
	'pptx.pdfExport.summary.simplified.one': 'auf 1 Folie wurden Effekte vereinfacht',
	'pptx.pdfExport.summary.simplified.other': 'auf {{count}} Folien wurden Effekte vereinfacht',
	'pptx.pdfExport.diagnostic.filterAsPicture':
		'Weiche Kanten, Weichzeichnen, Leuchten und andere Filtereffekte werden als Bilder gezeichnet.',
	'pptx.pdfExport.diagnostic.blendAsPicture': 'Mischmodi und Masken werden als Bilder gezeichnet.',
	'pptx.pdfExport.diagnostic.threeDAsPicture':
		'3D-Drehung und Perspektive werden als Bilder gezeichnet.',
	'pptx.pdfExport.diagnostic.mediaAsPicture':
		'Videos, 3D-Modelle und Canvas-Zeichnungen werden als Bilder ihres aktuellen Frames gezeichnet.',
	'pptx.pdfExport.diagnostic.textShadowAsPicture':
		'Textschatten werden als Bilder unter dem durchsuchbaren Text gezeichnet.',
	'pptx.pdfExport.diagnostic.svgAsPicture':
		'Zeichnungen, die dieser Export nicht als Vektorgrafik schreiben kann, werden als Bilder gezeichnet.',
	'pptx.pdfExport.diagnostic.fillTextAsPicture':
		'Text mit Farbverlaufs- oder Bildfüllung wird als Bild gezeichnet.',
	'pptx.pdfExport.diagnostic.boxFillAsPicture':
		'Füllungen, die dieser Export nicht als Vektorgrafik schreiben kann (Muster, Bildkacheln, radiale und halbtransparente Farbverläufe), werden als Bilder gezeichnet.',
	'pptx.pdfExport.diagnostic.borderAsPicture':
		'Rahmen mit anderen Stilen als durchgezogen, gestrichelt oder gepunktet oder mit unterschiedlichen Seiten an einem abgerundeten Feld werden als Bilder gezeichnet.',
	'pptx.pdfExport.diagnostic.shadowAsPicture': 'Schatten werden als Bilder gezeichnet.',
	'pptx.pdfExport.diagnostic.verticalTextAsPicture': 'Vertikaler Text wird als Bild gezeichnet.',
	'pptx.pdfExport.diagnostic.textColourAsPicture':
		'Text in einer Farbe, die dieser Export nicht schreiben kann, wird als Bild gezeichnet.',
	'pptx.pdfExport.diagnostic.complexScriptAsPicture':
		'Rechts-nach-links-Text und Schriften, die Zeichen verbinden oder kombinieren, werden als Bilder mit durchsuchbarem Text darüber gezeichnet.',
	'pptx.pdfExport.diagnostic.noFontTextAsPicture':
		'Text mit Zeichen, die keine Schriftart in der Datei darstellen kann (etwa Chinesisch, Japanisch oder Koreanisch, wenn die Präsentation keine Schriftart dafür einbettet), wird als Bild mit durchsuchbarem Text darüber gezeichnet.',
	'pptx.pdfExport.diagnostic.slideTooMuchText':
		'Eine Folie enthält mehr Text, als dieser Export liest, und wird daher als ein Bild gezeichnet.',
	'pptx.pdfExport.diagnostic.slideTooManyEffects':
		'Eine Folie hat zu viele Effekte, um sie einzeln zu zeichnen, und wird daher als ein Bild gezeichnet.',
	'pptx.pdfExport.diagnostic.slideNotCaptured':
		'Eine Folie konnte nicht erfasst werden und bleibt leer.',
	'pptx.pdfExport.diagnostic.backgroundPictureMissing':
		'Ein Hintergrundbild konnte nicht gelesen werden und wird weggelassen.',
	'pptx.pdfExport.diagnostic.pictureMissing':
		'Ein Bild konnte nicht gelesen werden und wird weggelassen.',
	'pptx.pdfExport.diagnostic.effectMissing':
		'Ein Effekt konnte nicht als Bild erfasst werden und wird weggelassen.',
	'pptx.pdfExport.diagnostic.drawingPartMissing':
		'Ein Teil einer Zeichnung konnte nicht als Vektorgrafik geschrieben werden und wird weggelassen.',
	'pptx.pdfExport.diagnostic.unsearchableText':
		'Mancher Text verwendet Zeichen, die keine verfügbare Schriftart enthält, und wird daher weggelassen.',
	'pptx.pdfExport.diagnostic.fontStandard':
		'Text in „{{family}}“ verwendet eine PDF-Standardschrift ({{substitute}}), weil die Schriftdatei nicht zum Einbetten verfügbar ist; die Positionen entsprechen der Folie, die Buchstabenformen weichen ab.',
	'pptx.pdfExport.diagnostic.fontStandardDefaultFamily':
		'Text in der Standardschriftart verwendet eine PDF-Standardschrift ({{substitute}}), weil die Schriftdatei nicht zum Einbetten verfügbar ist; die Positionen entsprechen der Folie, die Buchstabenformen weichen ab.',
	'pptx.pdfExport.diagnostic.fontSubstituted':
		'Text in „{{family}}“ verwendet stattdessen die eingebettete Schriftart „{{substitute}}“.',
	'pptx.pdfExport.diagnostic.fontSubstitutedDefaultFamily':
		'Text in der Standardschriftart verwendet stattdessen die eingebettete Schriftart „{{substitute}}“.',
	'pptx.pdfExport.diagnostic.searchableText':
		'Text in Bildern bleibt als unsichtbarer Text durchsuchbar.',
	'pptx.pdfExport.diagnostic.groupOpacity':
		'Die Transparenz einer Gruppe wird auf jedes ihrer Objekte einzeln angewendet.',
	'pptx.pdfExport.diagnostic.gradientOpacityAveraged':
		'Die Transparenz von Farbverläufen wird gemittelt: Eine PDF-Schattierung hat eine einzige Deckkraft.',
	'pptx.pdfExport.diagnostic.gradientSpreadApproximated':
		'Wiederholte und gespiegelte Farbverläufe werden als ein Farbverlauf gezeichnet.',
	'pptx.pdfExport.diagnostic.gradientFlatFill':
		'Ein Farbverlauf ohne nutzbare Fläche wird als eine einheitliche Farbe (der mittlere Punkt) gezeichnet.',
	'pptx.pdfExport.diagnostic.gradientOutlineFlat':
		'Konturen mit Farbverlauf werden als eine einheitliche Farbe (der mittlere Punkt) gezeichnet.',
	'pptx.pdfExport.diagnostic.shadowBlurDropped':
		'Die Weichzeichnung von Schatten wird nicht angewendet: Schatten werden scharfkantig gezeichnet.',
	'pptx.pdfExport.diagnostic.glowApproximated':
		'Leuchten wird durch einen scharfkantigen Rand mit halber Stärke angenähert (ohne Weichzeichnung).',
	'pptx.pdfExport.diagnostic.reflectionDropped':
		'Spiegelungen und maskierte Formen werden nicht gezeichnet.',
	'pptx.pdfExport.diagnostic.bevelDropped':
		'Abschrägungsprofile und Beleuchtung werden nicht gezeichnet.',
	'pptx.pdfExport.diagnostic.groupPerspectiveAffine':
		'Die Elemente einer perspektivisch gedrehten Gruppe verwenden eine affine Annäherung der Projektion.',
	'pptx.pdfExport.diagnostic.solid3dFlat':
		'Perspektive und Extrusion werden aus dem projizierten Umriss mit flacher Schattierung gezeichnet; Text und Bilder verwenden eine affine Annäherung der Projektion.',
	'pptx.pdfExport.diagnostic.perspectiveParallel':
		'3D-Drehung wird als Parallelprojektion ohne Perspektive gezeichnet.',
	'pptx.pdfExport.diagnostic.filterDropped':
		'Weichzeichnen und weiche Kanten werden nicht angewendet.',
};
