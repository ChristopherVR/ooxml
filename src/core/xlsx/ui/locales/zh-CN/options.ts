// File > Options strings: the shared Office option strings plus Excel's Formulas and Data panes.
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
};
