import type { DefaultTheme } from 'vitepress';

const sidebar: DefaultTheme.SidebarItem[] = [
	{
		text: '从这里开始',
		items: [
			{ text: '概览', link: '/zh/' },
			{ text: '快速开始', link: '/zh/getting-started' },
			{ text: '在线演示', link: '/demos' },
			{ text: '协作（在线演示）', link: '/zh/collaboration' },
			{ text: '架构', link: '/architecture' },
			{ text: '局限性', link: '/zh/limitations' },
		],
	},
	{
		text: '参考',
		items: [
			{ text: '元素 API 与事件', link: '/zh/api' },
			{ text: '主题定制', link: '/theming' },
		],
	},
	{
		text: '框架指南',
		items: [
			{ text: 'React', link: '/frameworks/react' },
			{ text: 'Vue', link: '/frameworks/vue' },
			{ text: 'Angular', link: '/frameworks/angular' },
			{ text: 'Svelte', link: '/frameworks/svelte' },
			{ text: 'Solid', link: '/frameworks/solid' },
			{ text: '原生 JS', link: '/frameworks/vanilla' },
		],
	},
	{
		text: '服务器',
		items: [
			{ text: '自带服务器', link: '/zh/server' },
			{ text: '部署参考服务器', link: '/zh/deploy' },
		],
	},
	{
		text: '演示',
		items: [
			{ text: '原生演示', link: '/demo/', target: '_self' },
			{ text: 'React 演示', link: '/demo-react/', target: '_self' },
			{ text: 'Vue 演示', link: '/demo-vue/', target: '_self' },
			{ text: 'Angular 演示', link: '/demo-angular/', target: '_self' },
			{ text: 'Svelte 演示', link: '/demo-svelte/', target: '_self' },
			{ text: 'Solid 演示', link: '/demo-solid/', target: '_self' },
			{ text: '关于演示', link: '/demos' },
		],
	},
	{
		text: '项目',
		items: [{ text: '发布软件包', link: '/releasing' }],
	},
];

export const zh = {
	label: '简体中文',
	lang: 'zh-CN',
	link: '/zh/',
	description:
		'开源、自带服务器的团队协作空间：频道与聊天、在线状态和 WebRTC 通话，提供 React、Vue、Angular、Svelte、Solid 和原生 JavaScript 绑定。',
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
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/teams/docs/:path',
			text: '在 GitHub 上编辑此页',
		},
		footer: {
			message:
				'基于 Apache-2.0 许可证发布。OpenTeams 不是 Microsoft Teams，与 Microsoft 没有任何关联。',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		sidebar,
		nav: [
			{
				text: '指南',
				link: '/zh/getting-started',
				activeMatch:
					'/zh/(getting-started|limitations|api|collaboration|server|deploy)|/(architecture|theming)',
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
