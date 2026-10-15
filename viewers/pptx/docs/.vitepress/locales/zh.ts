import type { DefaultTheme } from 'vitepress';

import { sidebar } from './zh-sidebar';

export const zh = {
	label: '简体中文',
	lang: 'zh-CN',
	link: '/zh/',
	description:
		'在浏览器和 Node.js 中解析、编辑、渲染和转换 PowerPoint 文件，提供 React、Vue 3、Angular、Svelte 和原生 JavaScript 组件。',
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
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/pptx/docs/:path',
			text: '在 GitHub 上编辑此页',
		},
		footer: {
			message: '基于 Apache-2.0 许可证发布',
			copyright: 'Copyright © 2025-present ChristopherVR',
		},
		sidebar,
		nav: [
			{ text: '开发指南', link: '/zh/guide/' },
			{ text: '使用指南', link: '/zh/user/' },
			{ text: '版本发布', link: '/zh/releases/' },
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
