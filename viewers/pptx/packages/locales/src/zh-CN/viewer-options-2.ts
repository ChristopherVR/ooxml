export const translations = {
	'pptx.options.addIns.mtxDecompressor': 'MicroType Express 字体解压缩器',
	'pptx.options.addIns.mtxDecompressorDescription': '解压缩嵌入的 MicroType Express 字体。',
	'pptx.options.addIns.collaboration': '实时协作',
	'pptx.options.addIns.collaborationDescription': '基于 CRDT 的共同编辑、在线状态和同步中继支持。',
	'pptx.options.addIns.locales': '语言包',
	'pptx.options.addIns.localesDescription': '为查看器界面提供更多显示语言。',
	'pptx.options.trust.label': '信任中心',
	'pptx.options.trust.description': '帮助确保您的文档安全可靠。',
	'pptx.options.trust.settingsTitle': '安全设置',
	'pptx.options.trust.settingsDescription':
		'这些设置有助于确保您打开的内容安全。我们建议使用默认值。',
	'pptx.options.trust.protectedView': '在受保护的视图中打开演示文稿',
	'pptx.options.trust.protectedViewInfo': '文件以只读方式打开；使用启用编辑进行更改。',
	'pptx.options.trust.allowExternalContent': '允许外部内容（远程图像和媒体）',
	'pptx.options.trust.allowExternalContentInfo': '关闭时，仅显示文件中嵌入的内容。',
	'pptx.options.trust.allowFontDownloads': '下载未安装或未嵌入的字体（Google Fonts）',
	'pptx.options.trust.allowFontDownloadsInfo':
		'开启时，浏览器会连接 fonts.googleapis.com 和 fonts.gstatic.com（或贵组织设置的字体服务器）以获取缺少的字体。关闭时，文本使用已安装、已嵌入或替代字体，因此换行可能与 PowerPoint 不同。',
	'pptx.options.trust.confirmHyperlinks': '打开外部超链接之前确认',
} as const;
