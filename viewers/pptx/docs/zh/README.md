# 维护文档

文档站点使用 VitePress。它的依赖安装在此目录中，与根目录的 Bun 工作区相互独立。

## 本地运行

在仓库根目录执行：

```bash
bun run docs:build   # installs docs dependencies, syncs release pages, builds site
bun run docs:dev     # starts VitePress after that initial installation
bun run docs:preview # previews the built site
```

如果想在不先做生产构建的情况下开发，请先在 `docs/` 中运行 `bun install`，再回到根目录运行
`bun run docs:dev`。输出位于 `docs/.vitepress/dist/`。部署的 base 路径是 `/ooxml/pptx/`。

## 在哪里更新内容

| 内容           | 位置                                                                    | 核对依据                                             |
| -------------- | ----------------------------------------------------------------------- | ---------------------------------------------------- |
| 安装与包概览   | 根目录及各包的 `README.md`                                              | 包清单、导出、对等依赖、构建配置                     |
| 开发者概念     | `guide/`                                                                | 核心类型、共享实现、演示配置                         |
| 无头 API       | `core/`                                                                 | `src/core/pptx/index.ts`、公开类、构建器与转换器类型 |
| 框架绑定 API   | `react/`、`vue/`、`angular/`、`svelte/`、`vanilla/`                     | 各绑定导出的属性、句柄、回调与默认值                 |
| 最终用户工作流 | `user/`                                                                 | 共享命令以及五个绑定实际的界面接线                   |
| 翻译           | `fr/`、`es/`、`de/`、`zh/`                                              | 对应的英文页面与当前实现                             |
| 首页文案与示例 | `.vitepress/theme/landing/copy/` 与 `code/samples.ts`                   | 与快速上手指南相同的公开 API                         |
| 导航           | `.vitepress/config.ts`                                                  | 现有页面与分区名称                                   |
| OpenXML 清单   | `architecture/openxml-conformance.md`                                   | 模式数据、解析与保存代码、符合性测试                 |
| 动画研究       | `../../../src/ui/src/pptx/render/animation-ppt-formula-ground-truth.md` | 公式求值器、播放集成与回归测试                       |

法语、西班牙语和德语指南只覆盖英文站点的一部分。简体中文（`zh/`）覆盖首页、开发者与用户指南、
核心以及五个框架绑定的 API 参考、MCP 工具、OpenXML 符合性参考和版本发布概览。自动生成的历史发布说明
保持原文语言，并从中文版本发布概览链接过去。`packages/locales` 中的界面字典是一个独立的系统；
新增界面语言不会自动翻译本站点。

请让翻译页面与英文页面保持一致。保留代码示例和 API 标识符，并使用与英文页面一致的显式标题 ID，
以便跨页锚点和语言切换继续有效。只要存在译文，中文内部链接就应指向 `/zh/` 页面。首页文案位于
`.vitepress/theme/landing/copy/zh.ts`；导航、侧边栏和搜索标签位于 `.vitepress/locales/zh.ts` 与
`zh-sidebar.ts`。

## 检查示例与表述

请对照公开入口检查示例，而不只是内部实现。部分 Hook 和服务仅通过 `/internals` 导出；不要暗示它们
是稳定的根导入。要包含必需的 CSS 导入、有尺寸的查看器容器，以及文件加载步骤，或明确说明所提供的输入。

核心加载器接受 `ArrayBuffer`。对于 Node 的 `Buffer`，请使用 `Uint8Array.from(buffer).buffer`
来获得恰好是文件字节的缓冲区；池化的 `buffer.buffer` 可能包含文件之外的字节。编辑富文本时，请检查保存
流程如何协调 `text` 与 `textSegments`；更改文字格式时要显式更新这些分段。

请区分解析、保留、编辑、播放和导出。识别某个 OOXML 结构并不等于完整的渲染保真度。请记录已知的近似
之处，并链接到限制或符合性页面，而不要声称普遍兼容。优先使用有源码依据的描述，而不是会变化的模块数、
测试数或工具数。

编辑后运行 `bun run docs:build`，对照类型或小型运行时用例检查修改过的示例，并运行
`git diff --check`。站点当前的 `ignoreDeadLinks` 设置意味着构建成功并不能证明每个链接都有效：
请把本地链接和导航作为单独的检查项来核对。

## 发布历史

`sync-changelogs.mjs` 会在开发和生产构建之前，把各包的变更日志复制到被忽略的 `releases/*.md` 页面中。
只有 `releases/index.md` 是手工维护的。请保持历史变更日志完整；发布工作流会在前面追加新条目，旧标签
可以被清理。不要根据剩余标签重新生成历史，也不要把生成的发布页面当作 API 文档来编辑。

当 `main` 上 `docs/` 下的内容变化时会运行 Pages 工作流，也可以手动触发或由发布工作流触发。它会构建文档和
演示应用。仅修改根目录 README 不会触发该路径过滤器。
