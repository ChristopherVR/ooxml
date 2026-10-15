import type { LandingCopy } from './types';

export const zh: LandingCopy = {
	hero: {
		titleTop: '团队聊天与通话，',
		titleAccent: '运行在你自己的服务器上。',
		sub: 'OpenTeams 是面向网页的团队协作空间：频道与聊天、在线状态、WebRTC 会议，以及在对话中共享的 Office 文件。一个 Web 组件，六种框架的轻量绑定，外加一个小型参考服务器。这是一个早期实现：不是 Microsoft Teams，也没有端到端加密。',
		start: { text: '快速开始', href: '/zh/getting-started' },
		demo: { text: '试用演示', href: '/demo/' },
		frameLabel: '本地模式下的 <teams-app>',
		tryLabel: '打开演示',
		imageAlt: 'OpenTeams 演示：频道列表、带回复和回应的对话，以及输入框',
	},
	liveDemo: {
		title: '就在这里试一试',
		note: '在线演示',
		introBefore:
			'这是在你的浏览器中运行的真实应用，背后没有服务器：聊天、在线状态和通话只在此浏览器的框架和标签页之间传输（BroadcastChannel）。切换到两个人，让 Ada 和 Bob 进入同一个房间，各自使用你选择的框架（例如 Ada 用 React，Bob 用 Vue），然后看着消息从一个窗格传到另一个窗格。实际使用需要服务器，请参阅',
		introLink: { text: '自带服务器', href: '/zh/server' },
		introAfter: '。',
		frameworkLabel: 'Ada 的框架',
		soloTab: '一个人',
		duoTab: '两个人',
		guestPicker: 'Bob 的框架',
		load: '加载在线演示',
		openFull: '打开完整应用',
		frameTitle: '{framework} OpenTeams 演示',
		hintDuo:
			'两个独立的应用，分别由各自的框架绑定构建，按名称加入同一个房间，并通过 BroadcastChannel 共享它。使用哪种框架并不重要：它们使用相同的核心协议。这是核心的本地模式，而不是托管服务：使用你自己的服务器时，这两个人可以位于不同的机器上。',
		hintSolo:
			'一切都在客户端运行，并保留在此浏览器中。可以打开完整应用，或切换到两个人，查看两位用户之间的在线状态和聊天。',
	},
	embed: {
		title: '两种嵌入方式',
		note: '演示应用',
		openDemo: '打开演示',
		demos: [
			{
				title: '原生：完整应用',
				copy: '一个 <teams-app> 元素：频道、带回复和回应的帖子、文件、搜索、在线状态和会议。在两个标签页中打开它，就能看到两个人在交谈。',
				href: '/demo/',
				link: '',
			},
			{
				title: 'React：组件与 hook',
				copy: '左侧是 <Teams /> 组件，右侧是基于 useTeams() hook 手写的面板，两者共用同一个工作空间。',
				href: '/demo-react/',
				link: '',
			},
		],
	},
	features: {
		title: '可以自行托管的协作空间',
		note: '功能一览',
		items: [
			{
				title: '频道与聊天',
				copy: '频道、带回复的帖子、回应、仅作者可用的编辑和删除、搜索和未读计数，全部位于同一个可合并并发编辑的 Yjs 文档中。',
				href: '/zh/getting-started',
				link: '快速开始',
				icon: 'chat',
			},
			{
				title: '在线状态',
				copy: '谁在线、输入提示，以及你自己设置的状态（可用、忙碌、离开）。在线状态只具参考性，与任何 Yjs awareness 一样。',
				href: '/architecture',
				link: '架构',
				icon: 'people',
			},
			{
				title: '基于 WebRTC 的会议',
				copy: '全网状通话中的音频、视频、屏幕共享和举手：没有媒体服务器，实际可靠地支持十来个人。TURN 由你自己运行。',
				href: '/zh/server#calls-stun-and-turn',
				link: '通话与 TURN',
				icon: 'video',
			},
			{
				title: '对话中的文件',
				copy: '附件会上传到你的服务器（或你自己的上传器），并以链接形式发布。Word、Excel、PowerPoint 和 Visio 文件会被识别，以便宿主打开它们。',
				href: '/zh/server#files',
				link: '文件存储',
				icon: 'file',
			},
			{
				title: '自带服务器',
				copy: '三个小型约定：y-websocket 同步、JSON 信令中继和可选的文件存储。openteams-server 在一个 Node 进程中完成这三项。',
				href: '/zh/server',
				link: '服务器指南',
				icon: 'server',
			},
			{
				title: '六种框架',
				copy: 'React、Vue、Angular、Svelte、Solid 和原生绑定。每个绑定都提供完整的组件，以及基于同一个客户端 store 的原始 hook。',
				href: '/frameworks/react',
				link: '框架指南',
				icon: 'code',
			},
		],
	},
	limits: {
		title: '它不是什么',
		note: '坦率的局限',
		items: [
			'没有端到端加密：你的服务器会存储并可以读取每一条消息和每一个文件。',
			'授权是服务器的职责；参考服务器只有一个共享令牌，没有用户账户。',
			'网状通话适合少数几个人，而不适合网络研讨会。',
			'没有一对一聊天、通知、日历、录制、背景模糊或实时字幕。',
			'不是 Microsoft Teams，与 Microsoft 没有任何关联，也不声称功能对等。',
		],
		all: { text: '全部局限性', href: '/zh/limitations' },
	},
	quickstart: {
		title: '几行代码即可拥有一个协作空间',
		note: '快速上手',
		introBefore: '安装适用于你的框架的绑定并渲染该组件。本地模式不需要服务器；实际使用时，请将 ',
		introAfter:
			' 指向你自己的同步和信令端点（参考服务器两者都提供）。Angular、Svelte 和 Solid 的选项相同。',
		fullGuide: '完整指南：{framework}',
	},
	faq: {
		title: '常见问题',
		note: 'FAQ',
		items: [
			{
				q: '可以免费用于商业用途吗？',
				a: '仓库采用 Apache-2.0 许可证。这些包以 openteams-*-viewer 和 openteams-server 的名称发布到 npm。',
			},
			{
				q: '需要服务器吗？',
				a: '实际使用需要：你需要自带服务器，用于同步（y-websocket）、信令和可选的文件存储。参考服务器 openteams-server 在一个 Node 进程中完成这三项。没有服务器时，演示以本地模式运行，只能连接同一浏览器的标签页。',
				href: '/zh/server',
				link: '自带服务器',
			},
			{
				q: '有端到端加密吗？',
				a: '没有。聊天以 Yjs 更新的形式经由你的服务器传输，服务器可以读取并存储它。通话媒体在对等端之间使用 DTLS-SRTP。',
				href: '/zh/limitations',
				link: '局限性',
			},
			{
				q: '这是 Microsoft Teams 吗？',
				a: '不是。OpenTeams 不是 Microsoft Teams，与 Microsoft 没有任何关联，也不声称功能对等。它提供频道、带回复的帖子、回应、文件、搜索、在线状态，以及带屏幕共享的会议。',
			},
			{
				q: '一次通话最多可以有多少人加入？',
				a: '通话是没有媒体服务器的全网状结构，因此适合少数几个人（默认上限为 12），而不适合网络研讨会。TURN 由你自己运行。',
				href: '/zh/server#calls-stun-and-turn',
				link: '通话与 TURN',
			},
			{
				q: '支持哪些框架？',
				a: 'React、Vue 3、Angular、Svelte 5、SolidJS 和纯 JavaScript。它们共用一个 <teams-app> 元素和一个核心客户端；每个绑定还提供原始 hook，方便你构建自己的界面。',
				href: '/frameworks/react',
				link: '框架指南',
			},
			{
				q: '它能处理 Word、Excel、PowerPoint 和 Visio 文件吗？',
				a: '文件以链接形式共享。Office 文件按名称识别，宿主应用可以在对应的查看器中打开它们；OpenTeams 本身不编辑这些文件。',
			},
		],
	},
	finale: {
		title: '你的协作空间，你的服务器。',
		introBefore: '安装一个绑定，用 ',
		introAfter:
			' 启动参考服务器，然后用两个标签页试试演示。采用 Apache-2.0 许可证，严格的 TypeScript，并且坦率说明不足之处。',
		start: { text: '快速开始', href: '/zh/getting-started' },
		github: '在 GitHub 上查看',
		columns: [
			{
				title: '产品',
				links: [
					{ text: '在线演示', href: '/demo/', external: false },
					{ text: '自带服务器', href: '/zh/server', external: false },
					{ text: '协作', href: '/zh/collaboration', external: false },
					{ text: '版本发布', href: '/releasing', external: false },
				],
			},
			{
				title: '文档',
				links: [
					{ text: '快速开始', href: '/zh/getting-started', external: false },
					{ text: '架构', href: '/architecture', external: false },
					{ text: '元素 API', href: '/zh/api', external: false },
					{ text: '局限性', href: '/zh/limitations', external: false },
				],
			},
			{
				title: '社区',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams',
						external: true,
					},
					{
						text: '问题反馈',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: '许可证',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/LICENSE',
						external: true,
					},
				],
			},
		],
		license: '© 2026 ChristopherVR · Apache-2.0',
		disclaimer: 'OpenTeams · 不是 Microsoft Teams，与 Microsoft 没有任何关联',
	},
};
