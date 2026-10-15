import type { LandingCopy } from './types';

export const zh: LandingCopy = {
	hero: {
		kicker: '早期实现 · Apache-2.0 · TypeScript',
		titleTop: '.docx 编辑，',
		titleAccent: '随时嵌入。',
		sub: '面向 React、Vue、Angular、Svelte、Solid 和原生 JavaScript 的浏览器 Word 编辑器。一个文档模型和一个基于 ProseMirror 的 <docx-editor> Web Component 承担全部工作；框架适配器只负责接入生命周期和事件。',
		start: { text: '快速开始', href: '/zh/getting-started' },
		demo: '在线演示',
		scroll: '向下滚动',
		frameCaption: '示例文档 document.docx · 在浏览器中实时运行',
		frameTry: '立即试用',
		frameAlt: 'docx-viewer 编辑器显示带有 Word 风格功能区的示例文档',
		copyLabel: '复制',
		copiedLabel: '已复制',
		entryLabel: '安装',
		notPublished: '已发布到 npm。每个框架包都是自包含的：内含编辑器，并自带 docx-core。',
	},
	features: {
		kicker: '功能',
		title: '目前已有的能力。',
		items: [
			{
				title: '统一的共享编辑器',
				copy: 'Word 风格的功能区、标题栏、文件后台视图和状态栏集成在一个 Web Component 中，支持撤销/重做、查找和替换、缩放以及只读模式。绑定层不添加任何自己的界面。',
				link: { text: '架构', href: '/zh/architecture' },
			},
			{
				title: '打印布局',
				copy: '分页视图支持换行、孤行/孤字控制与保持规则、分栏、表格行拆分、重复标题行、脚注定位以及打印。这是本编辑器自己的近似实现，并非 Word 的分页结果。',
				link: { text: '支持路线图', href: '/parity-roadmap' },
			},
			{
				title: '文档结构',
				copy: '样式、编号与项目符号列表、表格、节与分栏、页眉页脚、脚注，以及包含目录（TOC）在内的域。可插入内联文本框、页眉水印和页面边框。',
				link: { text: '待完成工作', href: '/outstanding-work' },
			},
			{
				title: '修订与批注',
				copy: '支持修订模式、接受与拒绝、上一处与下一处修订，以及带回复和解决功能的批注窗格。批注会写回 comments.xml。',
				link: { text: '编辑文本', href: '/zh/editing' },
			},
			{
				title: '谨慎保真',
				copy: '无改动的保存会返回原始字节。受支持的编辑会保留模型之外的包部件；会损坏不受支持内容的编辑将被拒绝，而不是被丢弃。',
				link: { text: '模型与保存', href: '/zh/architecture' },
			},
			{
				title: '旧版 .doc',
				copy: 'Word 97-2003 文件通过共享的 ole2 编解码器加载。可导入正文文本，对现有段落的受限编辑可保存回 .doc。格式和结构变更会被拒绝。',
				link: { text: '支持路线图', href: '/parity-roadmap' },
			},
			{
				title: '协作协议',
				copy: '与传输方式无关的协议，通过权威端交换经过验证的 ProseMirror 步骤，并提供临时的在线状态。网络、身份、权限和存储由你的应用负责。',
				link: { text: '协作', href: '/zh/collaboration' },
			},
			{
				title: '本地化',
				copy: '编辑器界面通过一个 locale 选项提供英语、法语、德语、西班牙语和简体中文。它只翻译界面文字，绝不翻译文档内容。',
				link: { text: '绑定', href: '/bindings' },
			},
		],
	},
	status: {
		kicker: '状态',
		title: '坦率说明尚未完成的部分。',
		copy: '这是一个早期实现，并不等同于 Microsoft Word：分页只是近似的，许多功能的呈现与 Word 并不一致，不支持受密码保护的文件，导出 DOCX 也无法恢复模型所省略的内容。支持路线图和待完成工作页面详细列出了这些差距。',
		link: { text: '支持路线图', href: '/parity-roadmap' },
		roadmapTitle: '路线图说明',
		roadmapCopy:
			'共享的 OOXML 逻辑正在迁移到私有的 ooxml-core 包，使本仓库最终只保留界面部分。该迁移正在进行中，尚不是已发布的功能。',
		roadmapLink: { text: 'ooxml-core 计划', href: '/ooxml-core-plan' },
	},
	quickstart: {
		kicker: '快速开始',
		title: '几行代码即可挂载编辑器。',
		copy: '每个适配器挂载的都是同一个 <docx-editor>。传入文档模型，监听变更，并为容器设置高度即可。',
		docsLabel: '框架指南',
		buildTitle: '从源码构建',
		buildCopy: '要运行演示或参与编辑器开发，请克隆本仓库。需要安装 Bun。',
		buildCommands: 'bun install\nbun run demo',
	},
	demos: {
		kicker: '在线演示',
		title: '就在这里试一试。',
		copy: '这是在你浏览器中运行的真实编辑器：针对每个框架适配器，由本仓库构建的演示应用实时嵌入于此。打开示例文档，或拆分视图，查看两个本地对等端共享同一份文档。',
		frameworkLabel: '框架',
		languageLabel: '语言',
		soloTab: '编辑器',
		collabTab: '协作',
		guestPicker: '窗口 B 的框架',
		load: '加载在线演示',
		loading: '正在加载在线编辑器',
		openFull: '打开完整应用',
		hostLabel: '主机（打开示例）',
		guestLabel: '按会话名称加入',
		soloHint:
			'在演示中使用“打开示例文档”，或拖入你自己的 DOCX 或 DOC 文件。切换框架会启动一个全新的编辑器。',
		collabHint:
			'两个窗口各自使用你选择的框架，按名称加入同一个会话。主机窗口运行内存中的参考权威端；另一个窗口接收文档和每一次编辑，并显示彼此的光标。会话基于 BroadcastChannel，因此只在此浏览器的窗口之间有效；没有网络后端，生产环境的传输由宿主应用负责。',
	},
	faq: {
		kicker: '常见问题',
		title: '常见问题解答。',
		items: [
			{
				q: '可以通过 npm 安装吗？',
				a: '可以。docx-core 以及每个框架一个的自包含编辑器包都已发布到 npm，例如 npm install docx-react-viewer。每个包独立版本管理。',
				link: { text: '发布策略', href: '/releasing' },
			},
			{
				q: '可以免费用于商业用途吗？',
				a: '本仓库采用 Apache-2.0 许可证。',
			},
			{
				q: '这与 Word 完全一致吗？',
				a: '不一致。分页只是近似的，许多功能的呈现与 Word 并不一致，导出 DOCX 会根据受支持的可见内容创建新文档。导入、布局、编辑、导出、保真和无障碍的对等程度分别单独跟踪。',
				link: { text: '支持路线图', href: '/parity-roadmap' },
			},
			{
				q: '编辑后的文件能在 Word 中重新打开吗？',
				a: '受支持的编辑会写回原始包，不安全的编辑会被拒绝，而不是悄悄降低文档质量。这是由回归测试检验的设计目标，并非无损导出的承诺。',
				link: { text: '架构', href: '/zh/architecture' },
			},
			{
				q: '支持哪些框架？',
				a: 'React、Vue 3、Angular、Svelte 5、SolidJS 和原生 JavaScript。它们共用同一个编辑器和模型；每个适配器只处理挂载、属性更新和事件转发。',
				link: { text: '绑定', href: '/bindings' },
			},
			{
				q: '协作需要服务器吗？',
				a: '该协议与传输方式无关，并附带内存中的参考权威端。网络、身份验证、持久化和权限由你的应用负责。',
				link: { text: '协作', href: '/zh/collaboration' },
			},
			{
				q: '能打开旧的 .doc 文件吗？',
				a: '可以，但有限制。可导入正文文本，对现有段落的受限编辑可保存回 .doc；格式和结构变更会被拒绝。',
			},
			{
				q: '受密码保护的文件呢？',
				a: '目前不支持。',
			},
		],
	},
	finale: {
		kicker: '开始使用',
		title: '.docx 进，.docx 出。',
		sub: '阅读架构说明，选择你的框架指南，并用你自己的文档试试演示。采用 Apache-2.0 许可证，使用严格的 TypeScript，并坦率说明不足之处。',
		quick: { text: '阅读指南', href: '/zh/getting-started' },
		github: '在 GitHub 上查看',
		columns: [
			{
				title: '产品',
				links: [
					{
						text: '在线演示',
						href: 'https://christophervr.github.io/ooxml/docx/demo/',
						external: true,
					},
					{ text: '框架绑定', href: '/bindings' },
					{ text: '协作', href: '/zh/collaboration' },
					{ text: '版本发布', href: '/releasing' },
				],
			},
			{
				title: '文档',
				links: [
					{ text: '快速开始', href: '/zh/getting-started' },
					{ text: '架构', href: '/zh/architecture' },
					{ text: '编辑文本', href: '/zh/editing' },
					{ text: '支持路线图', href: '/parity-roadmap' },
					{ text: '待完成工作', href: '/outstanding-work' },
				],
			},
			{
				title: '社区',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx',
						external: true,
					},
					{
						text: '问题反馈',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: '许可证',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'docx-viewer · 面向 Web 的早期 Word 编辑器',
	},
};
