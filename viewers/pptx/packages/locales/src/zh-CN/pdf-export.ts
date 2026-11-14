/** 文件 > 导出 > PDF：选项对话框、近似处理通知和诊断消息。 */
export const translations: Record<string, string> = {
	'pptx.compatibility.partSeparator': '；',
	'pptx.compatibility.details': '详细信息',
	'pptx.pdfExport.dialog.title': 'PDF 选项',
	'pptx.pdfExport.dialog.range': '范围',
	'pptx.pdfExport.dialog.all': '全部',
	'pptx.pdfExport.dialog.current': '当前幻灯片',
	'pptx.pdfExport.dialog.slides': '幻灯片',
	'pptx.pdfExport.dialog.from': '从：',
	'pptx.pdfExport.dialog.to': '到：',
	'pptx.pdfExport.dialog.includeHidden': '包括隐藏的幻灯片',
	'pptx.pdfExport.dialog.export': '导出',
	'pptx.pdfExport.dialog.rangeInvalid':
		'请输入 1 到 {{count}} 之间的幻灯片编号，且起始编号不大于结束编号。',
	'pptx.pdfExport.dialog.rangeEmpty':
		'此范围内没有可导出的幻灯片。请包括隐藏的幻灯片或选择其他幻灯片。',
	'pptx.pdfExport.summary.message': 'PDF 已保存，但有近似处理：{{parts}}。',
	'pptx.pdfExport.summary.pictures.one': '1 个元素已保存为图片',
	'pptx.pdfExport.summary.pictures.other': '{{count}} 个元素已保存为图片',
	'pptx.pdfExport.summary.slidesAsPictures.one': '1 张幻灯片已保存为单张图片',
	'pptx.pdfExport.summary.slidesAsPictures.other': '{{count}} 张幻灯片已分别保存为单张图片',
	'pptx.pdfExport.summary.fonts.one': '1 种字体已被替换',
	'pptx.pdfExport.summary.fonts.other': '{{count}} 种字体已被替换',
	'pptx.pdfExport.summary.omitted.one': '1 张幻灯片上的内容已省略',
	'pptx.pdfExport.summary.omitted.other': '{{count}} 张幻灯片上的内容已省略',
	'pptx.pdfExport.summary.simplified.one': '1 张幻灯片上的效果已简化',
	'pptx.pdfExport.summary.simplified.other': '{{count}} 张幻灯片上的效果已简化',
	'pptx.pdfExport.diagnostic.filterAsPicture': '柔化边缘、模糊、发光和其他滤镜效果绘制为图片。',
	'pptx.pdfExport.diagnostic.blendAsPicture': '混合模式和蒙版绘制为图片。',
	'pptx.pdfExport.diagnostic.threeDAsPicture': '三维旋转和透视绘制为图片。',
	'pptx.pdfExport.diagnostic.mediaAsPicture': '视频、三维模型和画布绘图绘制为其当前帧的图片。',
	'pptx.pdfExport.diagnostic.textShadowAsPicture': '文字阴影绘制为可搜索文字下方的图片。',
	'pptx.pdfExport.diagnostic.svgAsPicture': '此导出无法写为矢量图形的绘图将绘制为图片。',
	'pptx.pdfExport.diagnostic.fillTextAsPicture': '使用渐变或图片填充的文字绘制为图片。',
	'pptx.pdfExport.diagnostic.boxFillAsPicture':
		'此导出无法写为矢量图形的填充（图案、图片平铺、径向渐变和半透明渐变）绘制为图片。',
	'pptx.pdfExport.diagnostic.borderAsPicture':
		'实线、虚线或点线以外样式的边框，或圆角框上各边不同的边框，绘制为图片。',
	'pptx.pdfExport.diagnostic.shadowAsPicture': '阴影绘制为图片。',
	'pptx.pdfExport.diagnostic.verticalTextAsPicture': '竖排文字绘制为图片。',
	'pptx.pdfExport.diagnostic.textColourAsPicture': '使用此导出无法写入的颜色的文字绘制为图片。',
	'pptx.pdfExport.diagnostic.complexScriptAsPicture':
		'从右到左的文字以及连接或组合字符的文字绘制为图片，其上覆盖可搜索文字。',
	'pptx.pdfExport.diagnostic.slideTooMuchText':
		'某张幻灯片的文字超出此导出可读取的数量，因此绘制为一张图片。',
	'pptx.pdfExport.diagnostic.slideTooManyEffects':
		'某张幻灯片的效果过多，无法分别绘制，因此绘制为一张图片。',
	'pptx.pdfExport.diagnostic.slideNotCaptured': '某张幻灯片无法捕获，保留为空白。',
	'pptx.pdfExport.diagnostic.backgroundPictureMissing': '某张背景图片无法读取，已省略。',
	'pptx.pdfExport.diagnostic.pictureMissing': '某张图片无法读取，已省略。',
	'pptx.pdfExport.diagnostic.effectMissing': '某个效果无法捕获为图片，已省略。',
	'pptx.pdfExport.diagnostic.drawingPartMissing': '绘图的一部分无法写为矢量图形，已省略。',
	'pptx.pdfExport.diagnostic.unsearchableText':
		'部分文字使用了任何可用字体都不包含的字符，因此已省略。',
	'pptx.pdfExport.diagnostic.fontStandard':
		'“{{family}}”中的文字使用标准 PDF 字体（{{substitute}}），因为字体文件无法嵌入；位置与幻灯片一致，字形有所不同。',
	'pptx.pdfExport.diagnostic.fontStandardDefaultFamily':
		'默认字体中的文字使用标准 PDF 字体（{{substitute}}），因为字体文件无法嵌入；位置与幻灯片一致，字形有所不同。',
	'pptx.pdfExport.diagnostic.fontSubstituted': '“{{family}}”中的文字改用嵌入字体“{{substitute}}”。',
	'pptx.pdfExport.diagnostic.fontSubstitutedDefaultFamily':
		'默认字体中的文字改用嵌入字体“{{substitute}}”。',
	'pptx.pdfExport.diagnostic.searchableText': '图片中的文字作为不可见文字保持可搜索。',
	'pptx.pdfExport.diagnostic.groupOpacity': '组的透明度分别应用于组中的每个对象。',
	'pptx.pdfExport.diagnostic.gradientOpacityAveraged':
		'渐变透明度取平均值：PDF 着色只有一个不透明度。',
	'pptx.pdfExport.diagnostic.gradientSpreadApproximated': '重复和反射渐变绘制为单个渐变。',
	'pptx.pdfExport.diagnostic.gradientFlatFill': '没有可用区域的渐变绘制为一种纯色（中间色标）。',
	'pptx.pdfExport.diagnostic.gradientOutlineFlat': '渐变轮廓绘制为一种纯色（中间色标）。',
	'pptx.pdfExport.diagnostic.shadowBlurDropped': '不应用阴影模糊：阴影以硬边绘制。',
	'pptx.pdfExport.diagnostic.glowApproximated': '发光以半强度的硬边光晕近似表示（无模糊）。',
	'pptx.pdfExport.diagnostic.reflectionDropped': '不绘制映像和蒙版形状。',
	'pptx.pdfExport.diagnostic.bevelDropped': '不绘制棱台轮廓和光照。',
	'pptx.pdfExport.diagnostic.groupPerspectiveAffine': '透视旋转组中的子对象使用投影的仿射近似。',
	'pptx.pdfExport.diagnostic.solid3dFlat':
		'透视和拉伸根据投影轮廓以平面着色绘制；文字和图片使用投影的仿射近似。',
	'pptx.pdfExport.diagnostic.perspectiveParallel': '三维旋转以无透视的平行投影绘制。',
	'pptx.pdfExport.diagnostic.filterDropped': '不应用模糊和柔化边缘滤镜。',
};
