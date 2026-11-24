// File > Options strings: the shared Office option strings plus Excel's Formulas, Data and Trust
// Center panes.
import { optionStrings } from '../../../../i18n/index';

export const options: Record<string, string> = {
	...optionStrings['zh-CN'],
	'Change options related to formula calculation, performance, and error handling.':
		'更改与公式计算、性能和错误处理相关的选项。',
	'Calculation options': '计算选项',
	'Working with formulas': '使用公式',
	'R1C1 reference style': 'R1C1 引用样式',
	'Not available yet. Only A1 references are supported.': '暂不可用。仅支持 A1 引用。',
	'Change settings for importing and analyzing data.': '更改导入和分析数据的设置。',
	'Not available yet. Data import and query settings are not supported.':
		'暂不可用。不支持数据导入和查询设置。',
	// Trust Center (the wording of PowerPoint's, naming Excel).
	'Help keep your documents safe and secure.': '帮助确保您的文档安全可靠。',
	'Security settings': '安全设置',
	'These settings help keep the content you open safe. We recommend the defaults.':
		'这些设置有助于确保您打开的内容安全。我们建议使用默认值。',
	'Download fonts that are not installed or embedded (Google Fonts)':
		'下载未安装或未嵌入的字体（Google Fonts）',
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Excel.':
		'开启时，浏览器会连接 fonts.googleapis.com 和 fonts.gstatic.com（或贵组织设置的字体服务器）以获取缺少的字体。关闭时，文本使用已安装、已嵌入或替代字体，因此换行可能与 Excel 不同。',
	'This setting is managed by your organization': '此设置由你的组织管理',
	// Advanced > Edit Custom Lists
	'Edit Custom Lists': '编辑自定义列表',
	'Create lists for use in sorts and fill sequences.': '创建用于排序和填充序列的列表。',
	'Custom lists:': '自定义序列:',
	'List entries:': '输入序列:',
	'Press Enter to separate list entries.': '按 Enter 键分隔列表项。',
	'NEW LIST': '新序列',
	'Import list from cells:': '从单元格中导入序列:',
	Import: '导入',
	'Enter at least one list entry.': '请至少输入一个列表项。',
	'Enter a valid cell range such as A1:A5.': '请输入有效的单元格区域，例如 A1:A5。',
	'The selected range has no text or number entries to import.':
		'所选区域没有可导入的文本或数字项。',
};
