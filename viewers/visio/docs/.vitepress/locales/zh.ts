import type { DefaultTheme } from 'vitepress';

export const zh = {
	label: '简体中文',
	lang: 'zh-CN',
	link: '/zh/',
	description:
		'浏览器中的 Visio 图表查看器：一个共享的 Web 组件、一个 SVG 渲染器，以及面向 React、Vue、Angular、Svelte、Solid 和原生 JavaScript 的轻量适配器。',
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
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/visio/docs/:path',
			text: '在 GitHub 上编辑此页',
		},
		footer: {
			message: '基于 Apache-2.0 许可证发布。',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		nav: [
			{
				text: '指南',
				link: '/zh/getting-started',
				activeMatch:
					'/(zh/(getting-started|architecture|api|collaboration)|(getting-started|architecture|bindings|api|theming|collaboration|demos|frameworks))',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{ text: '版本发布', link: '/releasing', activeMatch: '/releasing' },
		],
		sidebar: [
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
					{ text: '查看器 API', link: '/zh/api' },
					{ text: '主题（英文）', link: '/theming' },
				],
			},
			{
				text: '使用查看器',
				items: [{ text: '协作（共享）', link: '/zh/collaboration' }],
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
					{ text: '功能清单（英文）', link: '/parity' },
					{ text: '验证记录（英文）', link: '/verification' },
					{ text: '语料库检查（英文）', link: '/corpus' },
					{ text: '套件外观来源（英文）', link: '/suite-provenance' },
					{ text: '包发布（英文）', link: '/releasing' },
				],
			},
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
