export const translations = {
	'pptx.compatibility.partSeparator': '、',
	'pptx.compatibility.details': '詳細',
	'pptx.pdfExport.dialog.title': 'PDF のオプション',
	'pptx.pdfExport.dialog.range': '範囲',
	'pptx.pdfExport.dialog.all': 'すべて',
	'pptx.pdfExport.dialog.current': '現在のスライド',
	'pptx.pdfExport.dialog.slides': 'スライド指定',
	'pptx.pdfExport.dialog.from': '開始:',
	'pptx.pdfExport.dialog.to': '終了:',
	'pptx.pdfExport.dialog.includeHidden': '非表示スライドを含める',
	'pptx.pdfExport.dialog.export': 'エクスポート',
	'pptx.pdfExport.dialog.rangeInvalid':
		'1 から {{count}} までのスライド番号を、開始が終了以下になるように入力してください。',
	'pptx.pdfExport.dialog.rangeEmpty':
		'この範囲にはエクスポートするスライドがありません。非表示スライドを含めるか、別のスライドを選択してください。',
	'pptx.pdfExport.summary.message': 'PDF は近似表現を含めて保存されました: {{parts}}。',
	'pptx.pdfExport.summary.pictures.one': '1 個の要素が画像として保存されました',
	'pptx.pdfExport.summary.pictures.other': '{{count}} 個の要素が画像として保存されました',
	'pptx.pdfExport.summary.slidesAsPictures.one': '1 枚のスライドが 1 枚の画像として保存されました',
	'pptx.pdfExport.summary.slidesAsPictures.other':
		'{{count}} 枚のスライドがそれぞれ 1 枚の画像として保存されました',
	'pptx.pdfExport.summary.fonts.one': '1 個のフォントが置き換えられました',
	'pptx.pdfExport.summary.fonts.other': '{{count}} 個のフォントが置き換えられました',
	'pptx.pdfExport.summary.omitted.one': '1 枚のスライドでコンテンツが省略されました',
	'pptx.pdfExport.summary.omitted.other': '{{count}} 枚のスライドでコンテンツが省略されました',
	'pptx.pdfExport.summary.simplified.one': '1 枚のスライドで効果が簡略化されました',
	'pptx.pdfExport.summary.simplified.other': '{{count}} 枚のスライドで効果が簡略化されました',
	'pptx.pdfExport.diagnostic.filterAsPicture':
		'ぼかし、ブラー、光彩などのフィルター効果は画像として描画されます。',
	'pptx.pdfExport.diagnostic.blendAsPicture': 'ブレンド モードとマスクは画像として描画されます。',
	'pptx.pdfExport.diagnostic.threeDAsPicture': '3-D 回転と遠近法は画像として描画されます。',
	'pptx.pdfExport.diagnostic.mediaAsPicture':
		'ビデオ、3D モデル、キャンバス描画は、現在のフレームの画像として描画されます。',
	'pptx.pdfExport.diagnostic.textShadowAsPicture':
		'文字の影は、検索可能なテキストの下に画像として描画されます。',
	'pptx.pdfExport.diagnostic.svgAsPicture':
		'このエクスポートでベクター グラフィックとして書き出せない描画は、画像として描画されます。',
	'pptx.pdfExport.diagnostic.fillTextAsPicture':
		'グラデーションまたは図で塗りつぶされたテキストは画像として描画されます。',
	'pptx.pdfExport.diagnostic.boxFillAsPicture':
		'このエクスポートでベクター グラフィックとして書き出せない塗りつぶし (パターン、図のタイル、放射状および半透明のグラデーション) は、画像として描画されます。',
	'pptx.pdfExport.diagnostic.borderAsPicture':
		'実線、破線、点線以外のスタイルの罫線や、角丸のボックスで辺ごとに異なる罫線は、画像として描画されます。',
	'pptx.pdfExport.diagnostic.shadowAsPicture': '影は画像として描画されます。',
	'pptx.pdfExport.diagnostic.verticalTextAsPicture': '縦書きテキストは画像として描画されます。',
	'pptx.pdfExport.diagnostic.textColourAsPicture':
		'このエクスポートで書き出せない色のテキストは画像として描画されます。',
	'pptx.pdfExport.diagnostic.complexScriptAsPicture':
		'右から左に書くテキストや、文字を連結または合成する文字体系は画像として描画され、その上に検索可能なテキストが重ねられます。',
	'pptx.pdfExport.diagnostic.noFontTextAsPicture':
		'ファイル内のどのフォントでも表示できない文字のテキスト (プレゼンテーションに中国語、日本語、韓国語のフォントが埋め込まれていない場合など) は画像として描画され、その上に検索可能なテキストが重ねられます。',
	'pptx.pdfExport.diagnostic.slideTooMuchText':
		'スライドのテキストがこのエクスポートで読み取れる量を超えているため、1 枚の画像として描画されます。',
	'pptx.pdfExport.diagnostic.slideTooManyEffects':
		'スライドの効果が多すぎて個別に描画できないため、1 枚の画像として描画されます。',
	'pptx.pdfExport.diagnostic.slideNotCaptured':
		'スライドをキャプチャできなかったため、空白のままになっています。',
	'pptx.pdfExport.diagnostic.backgroundPictureMissing':
		'背景の図を読み取れなかったため、省略されています。',
	'pptx.pdfExport.diagnostic.pictureMissing': '図を読み取れなかったため、省略されています。',
	'pptx.pdfExport.diagnostic.effectMissing':
		'効果を画像としてキャプチャできなかったため、省略されています。',
	'pptx.pdfExport.diagnostic.drawingPartMissing':
		'描画の一部をベクター グラフィックとして書き出せなかったため、省略されています。',
	'pptx.pdfExport.diagnostic.unsearchableText':
		'一部のテキストに、利用可能などのフォントにもない文字が使われているため、省略されています。',
	'pptx.pdfExport.diagnostic.fontStandard':
		'フォント ファイルを埋め込めないため、「{{family}}」のテキストには PDF 標準フォント ({{substitute}}) が使用されます。位置はスライドと一致しますが、字形は異なります。',
	'pptx.pdfExport.diagnostic.fontStandardDefaultFamily':
		'フォント ファイルを埋め込めないため、既定のフォントのテキストには PDF 標準フォント ({{substitute}}) が使用されます。位置はスライドと一致しますが、字形は異なります。',
	'pptx.pdfExport.diagnostic.fontSubstituted':
		'「{{family}}」のテキストには、代わりに埋め込みフォント「{{substitute}}」が使用されます。',
	'pptx.pdfExport.diagnostic.fontSubstitutedDefaultFamily':
		'既定のフォントのテキストには、代わりに埋め込みフォント「{{substitute}}」が使用されます。',
	'pptx.pdfExport.diagnostic.searchableText':
		'画像内のテキストは、非表示のテキストとして検索可能なまま残ります。',
	'pptx.pdfExport.diagnostic.groupOpacity':
		'グループに設定された透明度は、グループ内の各オブジェクトに個別に適用されます。',
	'pptx.pdfExport.diagnostic.gradientOpacityAveraged':
		'PDF のシェーディングは 1 つの不透明度しか持てないため、グラデーションの透明度は平均化されます。',
	'pptx.pdfExport.diagnostic.gradientSpreadApproximated':
		'繰り返しおよび反射のグラデーションは、1 つのグラデーションとして描画されます。',
	'pptx.pdfExport.diagnostic.gradientFlatFill':
		'有効な領域がないグラデーションは、1 色の単色 (中間の分岐点の色) で描画されます。',
	'pptx.pdfExport.diagnostic.gradientOutlineFlat':
		'グラデーションの枠線は、1 色の単色 (中間の分岐点の色) で描画されます。',
	'pptx.pdfExport.diagnostic.shadowBlurDropped':
		'影のぼかしは適用されず、影は境界のはっきりした状態で描画されます。',
	'pptx.pdfExport.diagnostic.glowApproximated':
		'光彩は、半分の強さの境界がはっきりした輪郭 (ぼかしなし) で近似されます。',
	'pptx.pdfExport.diagnostic.reflectionDropped': '反射とマスクされた図形は描画されません。',
	'pptx.pdfExport.diagnostic.bevelDropped': '面取りの形状と照明は描画されません。',
	'pptx.pdfExport.diagnostic.groupPerspectiveAffine':
		'遠近法で回転されたグループの子要素には、投影の 1 つのアフィン近似が使用されます。',
	'pptx.pdfExport.diagnostic.solid3dFlat':
		'遠近法と奥行きは、投影された輪郭からフラット シェーディングで描画されます。テキストと図には投影の 1 つのアフィン近似が使用されます。',
	'pptx.pdfExport.diagnostic.perspectiveParallel':
		'3-D 回転は、遠近法なしの平行投影として描画されます。',
	'pptx.pdfExport.diagnostic.filterDropped': 'ブラーとぼかしのフィルターは適用されません。',
} as const;
