# 在线演示

同一个演示应用会为每个框架适配器各构建一次，并与本文档一同发布。每个演示都是完全在浏览器中运行的真实 `<xlsx-editor>`：工作簿在标签页中解析、编辑和保存，不会上传任何内容，背后也没有服务器。

| 演示                                              | 适配器                             |
| ------------------------------------------------- | ---------------------------------- |
| [React](/demo/){target="_self"}                   | `@christophervr/xlsx-react-viewer` |
| [Vue](/demo-vue/){target="_self"}                 | `xlsx-vue-viewer`                  |
| [Angular](/demo-angular/){target="_self"}         | `xlsx-angular-viewer`              |
| [Svelte](/demo-svelte/){target="_self"}           | `xlsx-svelte-viewer`               |
| [Solid](/demo-solid/){target="_self"}             | `xlsx-solid-viewer`                |
| [原生 JavaScript](/demo-vanilla/){target="_self"} | `xlsx-vanilla-viewer`              |

首页通过框架切换器嵌入了这些演示：[打开在线演示部分](/zh/#live-demo)。

## 在演示中可以做什么 {#what-you-can-do-in-a-demo}

- 拖入 `.xlsx`、`.xlsm`、旧版 `.xls` 或 `.csv` 文件，新建工作簿，或打开示例工作簿。
- 添加 `?sample=1` 可立即打开示例工作簿，添加 `?locale=de`（任意受支持的标签）可以以其他界面语言启动。
- 使用演示中的主题按钮或本站的主题切换：演示会实时跟随 `vitepress-theme-appearance` 偏好设置（参见[主题](/zh/theming)）。

## 协作 {#collaboration}

演示没有单独的协作页面：协同编辑已内置于编辑器中。选择**共享**，然后开始一个会话；同一浏览器的另一个窗口或标签页可以通过邀请链接或会话名称加入。有关房间选项以及共享和不共享的内容，请参见[协作](/zh/collaboration)。

## 在此嵌入 {#embedded-here}

<iframe
	src="/ooxml/xlsx/demo/?sample=1"
	title="xlsx-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## 在本地运行 {#run-them-locally}

```bash
bun install   # at the repository root
cd viewers/xlsx
bun run demo
```

`bun run docs:build`（在 `docs/` 目录下运行）会与 Pages 工作流完全相同地构建文档以及全部六个演示。
