import type { DefaultTheme } from 'vitepress';

const sidebar: DefaultTheme.SidebarItem[] = [
	{
		text: '开始使用',
		items: [
			{ text: '概览', link: '/zh/' },
			{ text: '快速开始', link: '/zh/getting-started' },
			{ text: '在线演示（英文）', link: '/demos' },
			{ text: '架构', link: '/zh/architecture' },
			{ text: '框架绑定（英文）', link: '/bindings' },
		],
	},
	{
		text: '参考',
		items: [
			{ text: '元素 API', link: '/zh/api' },
			{ text: '主题（英文）', link: '/theming' },
			{ text: '本地化（英文）', link: '/localization' },
		],
	},
	{
		text: '框架指南',
		items: [
			{ text: 'React（英文）', link: '/frameworks/react' },
			{ text: 'Vue（英文）', link: '/frameworks/vue' },
			{ text: 'Angular（英文）', link: '/frameworks/angular' },
			{ text: '原生 JS（英文）', link: '/frameworks/vanilla' },
			{ text: 'Svelte（英文）', link: '/frameworks/svelte' },
			{ text: 'Solid（英文）', link: '/frameworks/solid' },
		],
	},
	{
		text: '项目状态',
		items: [
			{ text: '功能与限制', link: '/zh/features' },
			{ text: '协作', link: '/zh/collaboration' },
			{ text: '软件包发布（英文）', link: '/releasing' },
			{
				text: '更新日志（英文）',
				link: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/CHANGELOG.md',
			},
		],
	},
];

export const zh = {
	label: '简体中文',
	lang: 'zh-CN',
	link: '/zh/',
	description:
		'基于浏览器的 Excel 电子表格编辑器：一个工作簿模型、一个 Web 组件编辑器，并提供 React、Vue、Angular、Svelte、Solid 和原生 JavaScript 适配器。',
	themeConfig: {
		langMenuLabel: '切换语言',
		skipToContentLabel: '跳转到内容',
		darkModeSwitchLabel: '外观',
		lightModeSwitchTitle: '切换到浅色模式',
		darkModeSwitchTitle: '切换到深色模式',
		sidebarMenuLabel: '菜单',
		returnToTopLabel: '返回顶部',
		outline: { label: '本页目录' },
		docFooter: { prev: '上一页', next: '下一页' },
		lastUpdated: { text: '最后更新' },
		editLink: {
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/xlsx/docs/:path',
			text: '在 GitHub 上编辑此页',
		},
		footer: {
			message: '基于 Apache-2.0 许可证发布',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		sidebar,
		nav: [
			{
				text: '指南',
				link: '/zh/getting-started',
				activeMatch:
					'/(zh/)?(getting-started|architecture|bindings|api|theming|localization|demos|features|collaboration|frameworks)',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{ text: '版本发布', link: '/releasing', activeMatch: '/releasing' },
		],
		search: {
			provider: 'local',
			options: {
				translations: {
					button: { buttonText: '搜索', buttonAriaLabel: '搜索文档' },
					modal: {
						displayDetails: '显示详细内容',
						resetButtonTitle: '清空搜索',
						backButtonTitle: '关闭搜索',
						noResultsText: '未找到相关结果',
						footer: {
							selectText: '选择',
							selectKeyAriaLabel: '回车键',
							navigateText: '切换',
							navigateUpKeyAriaLabel: '向上方向键',
							navigateDownKeyAriaLabel: '向下方向键',
							closeText: '关闭',
							closeKeyAriaLabel: 'Esc 键',
						},
					},
				},
			},
		},
	} satisfies DefaultTheme.Config,
};
