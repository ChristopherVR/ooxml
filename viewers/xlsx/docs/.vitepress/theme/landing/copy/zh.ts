import type { LandingCopy } from './types';

export const zh: LandingCopy = {
	hero: {
		kicker: '早期实现 · Apache-2.0 · TypeScript',
		titleTop: '.xlsx 编辑，',
		titleAccent: '可嵌入任何应用。',
		sub: '面向 React、Vue、Angular、Svelte、Solid 和原生 JavaScript 的浏览器电子表格编辑器。一个带公式引擎的工作簿模型和一个 <xlsx-editor> Web 组件完成全部工作；框架适配器只负责接入生命周期和事件。',
		start: { text: '快速开始', href: '/zh/getting-started' },
		demo: '在线演示',
		scroll: '向下滚动',
		frameCaption: '示例 workbook.xlsx · 在浏览器中实时运行',
		frameTry: '立即试用',
		frameAlt: 'xlsx-viewer 网格示意图，包含功能区、编辑栏和工作表标签',
		copyLabel: '复制',
		copiedLabel: '已复制',
		entryLabel: '安装',
		notPublished: '已发布到 npm。每个框架软件包都是自包含的：它打包了编辑器，并自带工作簿模型。',
	},
	features: {
		kicker: '功能',
		title: '它的设计目标。',
		items: [
			{
				title: '一个共享的编辑器',
				copy: '功能区、名称框、编辑栏、网格、工作表标签和状态栏都在一个 Web 组件中，支持撤销和重做以及只读模式。绑定不添加任何自己的界面。',
				link: { text: '元素 API', href: '/zh/api' },
			},
			{
				title: '会重新计算的公式',
				copy: 'ooxml-core 中的公式引擎负责解析和计算单元格公式，每次编辑后重新计算依赖单元格。未知函数显示 #NAME?，而不是猜测结果。',
				link: { text: '功能与限制', href: '/zh/features' },
			},
			{
				title: '按保存时的样子呈现格式',
				copy: '字体、填充、边框、对齐、数字格式、合并单元格、冻结窗格、列宽和行高都从文件中读取，并由网格绘制。',
				link: { text: '功能与限制', href: '/zh/features' },
			},
			{
				title: '条件格式与数据验证',
				copy: '色阶、数据条、图标集和单元格规则基于实时数值计算；下拉列表验证会在网格中提供其选项。',
				link: { text: '功能与限制', href: '/zh/features' },
			},
			{
				title: '图表，仅用于显示',
				copy: '常见图表类型根据实时单元格值绘制为 SVG。不提供图表编辑；图表部件会原样写回。',
				link: { text: '功能与限制', href: '/zh/features' },
			},
			{
				title: '旧版 .xls 与 CSV',
				copy: 'Excel 97-2003 工作簿通过 ooxml-core 内的共享 ole2 编解码器加载，并保存为 .xlsx。CSV 可以打开，也可以从当前工作表导出。',
				link: { text: '功能与限制', href: '/zh/features' },
			},
			{
				title: '谨慎地保留内容',
				copy: '模型未覆盖的部分（VBA 项目、透视缓存、自定义 XML、图表细节）在保存时会随所在工作表一并带过。',
				link: { text: '架构', href: '/zh/architecture' },
			},
			{
				title: '本地化',
				copy: '界面通过一个 locale 选项支持英语、法语、德语、西班牙语和简体中文。工作簿内容绝不会被翻译。',
				link: { text: '本地化', href: '/localization' },
			},
		],
	},
	status: {
		kicker: '状态',
		title: '坦诚说明尚未完成的部分。',
		copy: '这是一个早期实现，并不等同于 Microsoft Excel。数据透视表、切片器、宏和外部链接会被保留，但不显示也不编辑；迷你图可显示但不可编辑；函数库很大但并不完整；不支持受密码保护的文件。功能页列出了这些差距。',
		link: { text: '功能与限制', href: '/zh/features' },
		roadmapTitle: '逻辑所在之处',
		roadmapCopy:
			'工作簿模型、公式引擎、解析器、序列化器和 .xls 读取器都位于公开的 ooxml-core 软件包的 xlsx 区域。本仓库只包含 UI。',
		roadmapLink: { text: '架构', href: '/zh/architecture' },
	},
	quickstart: {
		kicker: '快速开始',
		title: '几行代码即可挂载编辑器。',
		copy: '每个适配器挂载的都是同一个 <xlsx-editor>。传入工作簿或文件字节，监听变化，并给容器指定高度。',
		docsLabel: '框架指南',
		buildTitle: '从源码构建',
		buildCopy: '要运行演示或开发编辑器本身，请克隆仓库。需要 Bun。',
		buildCommands: 'bun install\nbun run demo',
	},
	demos: {
		kicker: '在线演示',
		title: '就在这里试用。',
		copy: '这是在你的浏览器中运行的编辑器：由本仓库为每个框架适配器构建的演示应用，已实时嵌入并打开了示例工作簿。',
		frameworkLabel: '框架',
		languageLabel: '语言',
		load: '加载在线演示',
		loading: '正在加载在线编辑器',
		openFull: '打开完整应用',
		hint: '示例包含三个工作表，带有公式、条件格式、验证列表和图表。打开完整应用可以放入你自己的 .xlsx、.xls 或 .csv 文件。切换框架会启动一个全新的编辑器。',
	},
	faq: {
		kicker: '常见问题',
		title: '常见问题解答。',
		items: [
			{
				q: '可以从 npm 安装吗？',
				a: '可以。每个框架都有一个自包含的编辑器软件包发布在 npm 上，例如 npm install @christophervr/xlsx-react-viewer。每个软件包独立版本化。',
				link: { text: '发布策略', href: '/releasing' },
			},
			{
				q: '可以免费用于商业用途吗？',
				a: '本仓库采用 Apache-2.0 许可证。',
			},
			{
				q: '这等同于 Excel 吗？',
				a: '不等同。网格绘制模型所理解的内容，公式引擎覆盖了大量但不完整的函数，还有一些工作簿功能只被保留，而不显示也不编辑。',
				link: { text: '功能与限制', href: '/zh/features' },
			},
			{
				q: '保存的文件能在 Excel 中打开吗？',
				a: '保存的文件被写成无需修复即可打开，往返由 ooxml-core 中的测试检查。这是设计目标，而不是对无损导出的承诺。',
				link: { text: '架构', href: '/zh/architecture' },
			},
			{
				q: '支持哪些框架？',
				a: 'React、Vue 3、Angular、Svelte 5、SolidJS 和原生 JavaScript。它们共享同一个编辑器和模型；每个适配器只处理挂载、属性更新和事件转发。',
				link: { text: '框架绑定', href: '/bindings' },
			},
			{
				q: '多人可以编辑同一个工作簿吗？',
				a: '目前还不行。电子表格尚未实现实时协作。',
				link: { text: '协作', href: '/zh/collaboration' },
			},
			{
				q: '可以打开旧的 .xls 文件吗？',
				a: '可以，只读：值、可解码的公式、基本格式、合并单元格和工作表布局都会加载，工作簿保存为 .xlsx。',
			},
			{
				q: '受密码保护的文件呢？',
				a: '目前不支持。',
			},
		],
	},
	finale: {
		kicker: '立即开始',
		title: '.xlsx 进，.xlsx 出。',
		sub: '阅读快速开始指南，选择你的框架，并用你自己的工作簿试用演示。采用 Apache-2.0 许可证，严格的 TypeScript，并坦诚说明差距。',
		quick: { text: '阅读指南', href: '/zh/getting-started' },
		github: '在 GitHub 上查看',
		columns: [
			{
				title: '产品',
				links: [
					{
						text: '在线演示',
						href: 'https://christophervr.github.io/ooxml/xlsx/demo/',
						external: true,
					},
					{ text: '框架绑定', href: '/bindings' },
					{ text: '元素 API', href: '/zh/api' },
					{ text: '版本发布', href: '/releasing' },
				],
			},
			{
				title: '文档',
				links: [
					{ text: '快速开始', href: '/zh/getting-started' },
					{ text: '在线演示', href: '/demos' },
					{ text: '架构', href: '/zh/architecture' },
					{ text: '主题', href: '/theming' },
					{ text: '功能与限制', href: '/zh/features' },
				],
			},
			{
				title: '社区',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx',
						external: true,
					},
					{
						text: '问题反馈',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: '许可证',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'xlsx-viewer · 面向 Web 的早期电子表格编辑器',
	},
};
