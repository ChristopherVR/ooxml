import type { LandingCopy } from './types';

export const zh: LandingCopy = {
	hero: {
		kicker: '公开测试版 · Apache-2.0 · TypeScript',
		titleTop: '.vsdx 查看，',
		titleAccent: '可嵌入。',
		sub: '将 Visio 图表带入你的 Web 应用。一个共享的 <visio-viewer> Web 组件和一个 SVG 渲染器，适用于 React、Vue、Angular、Svelte、Solid 和原生 JavaScript。浏览页面、检查形状，并进行受支持的纯文本编辑；框架适配器只负责连接生命周期和事件。',
		start: { text: '开始使用', href: '/zh/getting-started' },
		demo: '在线演示',
		scroll: '向下滚动',
		frameCaption: '示例工作流 · 在浏览器中实时运行',
		frameTry: '立即试用',
		frameAlt: 'visio-viewer 工作区，展示原创示例、页面导航、功能区和形状检查器',
		copyLabel: '复制',
		copiedLabel: '已复制',
		entryLabel: '安装',
		notPublished: '已发布到 npm。每个框架包都是自包含的：它打包了共享查看器、其适配器和 Worker。',
	},
	features: {
		kicker: '特性',
		title: '目前已有的功能。',
		items: [
			{
				title: '六种原生绑定',
				copy: 'React、Vue、Angular、Svelte、Solid 或原生 JavaScript。轻量的生命周期适配器将所有格式逻辑保留在 ooxml-core 和共享查看器中。',
				link: { text: '框架指南', href: '/frameworks/react' },
			},
			{
				title: '聚焦图表本身',
				copy: '切换页面、适应窗口和缩放、选择形状、搜索可见文本、调整受支持的图层可见性，并使用 Visio 风格的功能区、“文件”后台视图、“形状”窗口以及“平移和缩放”窗口。',
				link: { text: '在线演示', href: '/demos' },
			},
			{
				title: '你的浏览器，你的文件',
				copy: '打开本地 .vsdx 文件，并将当前页面导出为 SVG。没有文档上传，没有遥测，也不会远程加载文档。',
				link: { text: '快速开始', href: '/zh/getting-started' },
			},
			{
				title: '谨慎的编辑',
				copy: '实验性的、以源为依据的纯文本替换，范围有限的几何编辑，有上限的撤销/重做，以及明确的原始或已编辑 VSDX 副本。不受支持的内容会被拒绝，而不是被丢弃。',
				link: { text: '编辑限制', href: '/zh/getting-started#editing' },
			},
			{
				title: '旧版 .vsd 预览',
				copy: '二进制 VSD v11 文件通过 ooxml-core 中共享的 ole2 编解码器加载，作为只读预览，并带有经过诊断的回退样式。',
				link: { text: '功能清单', href: '/parity' },
			},
			{
				title: '实时共享',
				copy: '“文件 > 共享”会在同一浏览器的窗口之间通过 BroadcastChannel 启动 Yjs 会话。跨设备共享需要由宿主提供服务器传输，目前尚未构建。',
				link: { text: '协作', href: '/zh/collaboration' },
			},
			{
				title: '导出与打印准备',
				copy: '导出当前页面的 SVG，包含嵌入的栅格图像和可见的兼容性说明，并可通过 API 获得不可变的页面快照。不提供打印对话框、PDF 或位图导出。',
				link: { text: '查看器 API', href: '/zh/api' },
			},
			{
				title: '清楚标明的限制',
				copy: '每项功能都附有证据以及尚缺的部分。外观出色的功能区并不等于声称与 Visio 一致。',
				link: { text: '功能清单', href: '/parity' },
			},
		],
	},
	status: {
		kicker: '状态',
		title: '如实说明尚未完成的部分。',
		copy: '这是公开测试版，并非与 Microsoft Visio 完全一致。通用绘图、完整的 ShapeSheet 重新计算、富文本、组和大多数母版、PDF 导出以及在原生 Visio 中重新打开的保真度均未得到验证。生成的示例用于演示交互，并不能证明视觉保真度。',
		link: { text: '功能清单', href: '/parity' },
		roadmapTitle: '逻辑所在之处',
		roadmapCopy:
			'解析、ShapeSheet 解释、几何和编辑命令都位于 ooxml-core/visio。本仓库只包含 UI：一个控制器、一个 SVG 渲染器、一个元素和六个轻量适配器。',
		roadmapLink: { text: '架构', href: '/zh/architecture' },
	},
	quickstart: {
		kicker: '快速开始',
		title: '几行代码即可挂载查看器。',
		copy: '每个适配器都转发共享查看器相同的属性、事件和命令式句柄。为容器设置高度，然后传入 VisioDocument 或加载 File。',
		docsLabel: '框架指南',
		buildTitle: '从源码运行',
		buildCopy: '要运行演示环境或开发查看器本身，请克隆仓库。需要 Node.js 22.12 或更高版本。',
		buildCommands: 'npm ci --ignore-scripts\nnpm run dev',
	},
	demos: {
		kicker: '在线演示',
		title: '就在这里试用。',
		copy: '这是在你的浏览器中运行的真实查看器：由本仓库为每个框架适配器构建的演示环境，在此实时嵌入。打开示例绘图或本地 .vsdx；不会上传任何内容。',
		frameworkLabel: '框架',
		soloTab: '查看器',
		collabTab: '共享',
		guestPicker: '窗口 B 的框架',
		load: '加载在线演示',
		loading: '正在加载在线查看器',
		openFull: '打开完整应用',
		hostLabel: '窗口 A',
		guestLabel: '窗口 B',
		soloHint: '使用示例绘图，或打开你自己的 .vsdx 文件。切换框架会启动一个全新的查看器。',
		collabHint:
			'同一浏览器中的两个窗口，各自使用你选择的框架，并已按名称加入同一个“文件 > 共享”会话。窗口 A 打开示例，窗口 B 接收它。在其中一个窗口中编辑形状文本，另一个窗口中就会出现相应变化。它使用以会话名称为键的 BroadcastChannel；没有服务器，仅在此浏览器内有效，也不会上传任何内容。',
	},
	faq: {
		kicker: '常见问题',
		title: '常见问题解答。',
		items: [
			{
				q: '它能完全替代 Visio 吗？',
				a: '不能。这是一个早期实现。受支持的缓存几何、文本和图像通过同一条 SVG 路径渲染。复杂文档仍有缺口；功能清单将实现、证据和缺失的工作区分开列出。',
				link: { text: '功能清单', href: '/parity' },
			},
			{
				q: '可以从 npm 安装吗？',
				a: '可以。每个框架都有一个自包含的查看器包发布在 npm 上，例如 npm install visio-react-viewer。每个包独立进行版本管理。',
				link: { text: '快速开始', href: '/zh/getting-started' },
			},
			{
				q: '打开文件会上传它吗？',
				a: '不会。查看器读取你在浏览器中选择的 File。没有文档上传、遥测或远程文档获取。',
			},
			{
				q: '可以编辑并保存 .vsdx 文件吗？',
				a: '对于受支持的、以源为依据的形状，提供实验性的纯文本和范围有限的几何编辑，带有有上限的撤销/重做以及明确的 VSDX 副本下载。与母版关联的文本、富文本、字段、已签名的包和宏内容会被拒绝。源文件绝不会被覆盖，在原生 Visio 中重新打开仍未得到验证。',
				link: { text: '查看器 API', href: '/zh/api' },
			},
			{
				q: '可以导出或打印什么？',
				a: 'UI 可以导出当前页面的 SVG 以及原始或已编辑的 VSDX 副本。API 可以为将来的打印工作流准备不可变的页面快照；它不会打开打印对话框。PDF、位图导出和原生打印机布局均未实现。',
			},
			{
				q: '多人可以编辑同一个图表吗？',
				a: '仅限有限的方式：“文件 > 共享”通过 Yjs 会话连接同一浏览器的窗口。跨设备共享需要由宿主提供服务器传输，此版本不包含。',
				link: { text: '协作', href: '/zh/collaboration' },
			},
			{
				q: '示例能证明视觉保真度吗？',
				a: '不能。这个以编程方式创建的原创示例用于演示交互。生成的夹具并不能证明与 Microsoft Visio 在视觉上一致；请将重要的图纸与其在原生 Visio 中的渲染效果进行对比。',
			},
			{
				q: '支持哪些框架？',
				a: 'React、Vue 3、Angular、Svelte 5、SolidJS 和原生 JavaScript。它们共享同一个查看器元素；每个适配器只负责挂载、属性更新和事件转发。',
				link: { text: '框架指南', href: '/frameworks/react' },
			},
		],
	},
	finale: {
		kicker: '开始使用',
		title: '输入 .vsdx，得到图表。',
		sub: '阅读指南，选择你的框架，并用你自己的图纸试用演示环境。采用 Apache-2.0 许可，严格的 TypeScript，并如实说明其不足。',
		quick: { text: '阅读指南', href: '/zh/getting-started' },
		github: '在 GitHub 上查看',
		columns: [
			{
				title: '产品',
				links: [
					{
						text: '在线演示',
						href: 'https://christophervr.github.io/ooxml/visio/demo/',
						external: true,
					},
					{ text: '在线演示合集', href: '/demos' },
					{ text: '协作', href: '/zh/collaboration' },
					{ text: '版本发布', href: '/releasing' },
				],
			},
			{
				title: '文档',
				links: [
					{ text: '快速开始', href: '/zh/getting-started' },
					{ text: '架构', href: '/zh/architecture' },
					{ text: '查看器 API', href: '/zh/api' },
					{ text: '功能清单', href: '/parity' },
				],
			},
			{
				title: '社区',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio',
						external: true,
					},
					{
						text: '问题反馈',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: '许可证',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'visio-viewer · 面向 Web 的早期 Visio 查看器',
	},
};
