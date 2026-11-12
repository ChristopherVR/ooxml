# 在线演示

同一个演示环境会为每个框架适配器各构建一次，并与这些文档一同发布。每个演示都是真正的 `<visio-viewer>`，完全在你的浏览器中运行：图纸在本地解析和渲染，不会上传任何内容，背后也没有服务器。这是公开测试版；演示能够打开某个文件，并不能证明其与 Visio 的渲染保真度一致。

| 演示                                      | 适配器                 | 路由             |
| ----------------------------------------- | ---------------------- | ---------------- |
| [原生 JavaScript](/demo/){target="_self"} | `visio-vanilla-viewer` | `/demo/`         |
| [React](/demo-react/){target="_self"}     | `visio-react-viewer`   | `/demo-react/`   |
| [Vue](/demo-vue/){target="_self"}         | `visio-vue-viewer`     | `/demo-vue/`     |
| [Angular](/demo-angular/){target="_self"} | `visio-angular-viewer` | `/demo-angular/` |
| [Svelte](/demo-svelte/){target="_self"}   | `visio-svelte-viewer`  | `/demo-svelte/`  |
| [Solid](/demo-solid/){target="_self"}     | `visio-solid-viewer`   | `/demo-solid/`   |

原生 JavaScript 演示环境同时也提供于 [/demo-vanilla/](/demo-vanilla/){target="_self"}，这是其他查看器站点使用的名称。首页以框架切换器嵌入这些演示：[打开实时演示区域](/zh/#live-demo)。

## 在演示中可以做什么 {#what-you-can-do-in-a-demo}

- 打开原始示例图纸，或选择本地 `.vsdx` 文件。
- 在演示地址后添加 `?sample=1`，可直接打开示例；添加 `?embed=1` 则使用首页所用的紧凑嵌入布局。
- 使用演示中的主题按钮或本站的主题切换器：所有演示都会实时跟随 `vitepress-theme-appearance` 偏好设置（请参阅[主题](/zh/theming)）。

## 协作演示 {#collaboration-demo}

在任意演示地址后添加 `?share=<session>`，即可在加载时加入该共享会话。使用相同会话名称的两个不同框架演示会共享同一份图纸，例如在同一浏览器的两个标签页中分别打开 `/demo-vue/?sample=1&share=my-room` 和 `/demo-angular/?share=my-room`。首页会显示两个窗格，每个窗格各有一个框架选择器。支持与不支持的内容请参阅[协作](/zh/collaboration)。

## 嵌入在此处 {#embedded-here}

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1"
	title="visio-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## 在本地运行 {#run-them-locally}

```bash
bun install            # once, at the repository root
cd viewers/visio
bun run dev
```

演示源代码位于仓库根目录的 `demos/visio/` 中（`demo` 是共享的原生 JavaScript 演示环境，`demo-react`、`demo-vue`、`demo-angular`、`demo-svelte` 和 `demo-solid` 则挂载同一个工作区）。该查看器的脚本和 Vite 配置负责构建并提供这些演示。`bun run docs:build` 会按照仓库的 Pages 构建方式，构建文档和所有演示（仓库根目录的 `node scripts/build-pages.mjs` 会构建每个查看器的站点）。
