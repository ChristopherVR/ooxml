# 在线演示

同一个演示应用会为每个框架适配器各构建一次，并与本文档一同发布。每个演示都是运行在你浏览器中的真实 `<docx-editor>`：没有任何内容会被上传，其背后也没有服务器。

| 演示                                              | 适配器                | 源码                                                            |
| ------------------------------------------------- | --------------------- | --------------------------------------------------------------- |
| [React](/demo/){target="_self"}                   | `docx-react-viewer`   | `demos/docx/demo-vanilla` 使用 `VITE_DEMO_FRAMEWORK=react` 构建 |
| [Vue](/demo-vue/){target="_self"}                 | `docx-vue-viewer`     | 同一应用，`vue`                                                 |
| [Angular](/demo-angular/){target="_self"}         | `docx-angular-viewer` | 同一应用，`angular`                                             |
| [Svelte](/demo-svelte/){target="_self"}           | `docx-svelte-viewer`  | 同一应用，`svelte`                                              |
| [Solid](/demo-solid/){target="_self"}             | `docx-solid-viewer`   | 同一应用，`solid`                                               |
| [原生 JavaScript](/demo-vanilla/){target="_self"} | `docx-vanilla-viewer` | 同一应用，`vanilla`                                             |

首页通过框架切换器嵌入了这些演示：[打开在线演示部分](/zh/#live-demo)。

## 在演示中可以做什么 {#what-you-can-do-in-a-demo}

- 选择 **打开示例文档**（Open the sample document），或拖入你自己的 DOCX 或旧版 DOC 文件。文件只保留在当前标签页中。
- 在演示 URL 后添加 `?sample=1`，可直接打开示例文档；添加 `?locale=de`（也可用 `fr`、`es`、`zh-CN`），可以以其他界面语言启动。
- 使用演示中的主题按钮，或本站的主题切换：所有演示都会实时跟随 `vitepress-theme-appearance` 的偏好设置（参见[主题](/zh/theming)）。

## 协作演示 {#collaboration-demo}

在演示 URL 后添加 `?room=<session>`，即可加入同一浏览器中的共享会话。同时带有 `sample=1` 的窗口会作为主机；其他窗口通过名称加入，框架不限。例如在两个标签页中分别打开 `/demo-vue/?sample=1&room=my-room` 和 `/demo-solid/?room=my-room`。首页会显示两个窗格，每个窗格都有一个框架选择器。[双端页面](/demo/collaboration.html){target="_self"}还添加了一个“暂停投递”（Pause delivery）按钮，用于测试并发编辑。关于支持与不支持的内容，参见[协作](/zh/collaboration)。

## 此处嵌入的演示 {#embedded-here}

<iframe
	src="/ooxml/docx/demo/?sample=1"
	title="docx-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## 在本地运行 {#run-them-locally}

```bash
bun install          # from the repository root
cd viewers/docx
bun run demo
```

在 `docs/` 目录下运行 `bun run docs:build`，即可与 Pages 工作流一样构建文档以及全部六个演示。
