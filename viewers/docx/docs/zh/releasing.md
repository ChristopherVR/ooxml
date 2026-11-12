# 发布软件包

Word 查看器软件包由 `ooxml` 仓库自身的流程发布，与 `ooxml-core` 和 `ooxml-ui` 一同进行。`viewers/docx` 中没有单独的发布工作流。完整流程（规划器、每小时运行的 `release.yml`、可信发布、变更日志以及一次性设置）记录在仓库的[发布指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md)中。切勿手动运行 `npm publish` 或推送发布标签。

## 本查看器发布的软件包 {#what-this-viewer-publishes}

每个软件包都有自己的版本线以及自己的 `<npm-name>@<version>` git 标签。目录与 npm 名称之间的对应关系，记录在仓库根目录的 `scripts/viewer-packages.mjs` 中。

| 软件包                | 目录               | 说明                                                      |
| --------------------- | ------------------ | --------------------------------------------------------- |
| `docx-core`           | `packages/core`    | 对 `ooxml-core/docx` 与 `/docx/embedded` 的轻量重新导出。 |
| `docx-react-viewer`   | `packages/react`   | React 入口（`WordEditor`）。                              |
| `docx-vue-viewer`     | `packages/vue`     | Vue 3 入口。                                              |
| `docx-angular-viewer` | `packages/angular` | Angular 入口（`WordEditorComponent`）。                   |
| `docx-svelte-viewer`  | `packages/svelte`  | Svelte 5 入口。                                           |
| `docx-solid-viewer`   | `packages/solid`   | Solid 入口。                                              |
| `docx-vanilla-viewer` | `packages/vanilla` | 普通的 `<docx-editor>` 注册以及 `mountEditor`。           |
| `docx-viewer-mcp`     | `mcp`              | Word MCP 服务器（参见其 README 和根目录的 `mcp.md`）。    |

`web-component` 和 `bindings` 设置了 `private: true`，从不发布。`scripts/build-packages.mjs` 会将它们打包进每个框架包，因此一个 tarball 只会导入 `docx-core`、`ooxml-core`、`ooxml-ui`、ProseMirror 相关库及其框架 peer。`ooxml-ui` 和 `docx-core` 仍是真正的注册表依赖（不会被打包），而来自 `ole2` 的旧版 `.doc` 编解码器则内联在 `ooxml-core` 中，因此这里的任何内容都不依赖 `ole2`。

## 什么会触发发布 {#what-triggers-a-release}

根目录的规划器（`scripts/release-plan.mjs`，使用 `scripts/viewer-packages.mjs` 中的表格）会查看每个软件包自上次标签以来的约定式提交：

- `packages/<name>/` 下的非测试更改会发布该软件包。
- `packages/web-component` 或 `packages/bindings` 下的非测试更改会发布每个框架包，因为每个框架包都会包含该代码。`scripts/build-packages.mjs` 和 `tsconfig.release.json` 会发布本查看器中的所有内容。
- 发布 `docx-core` 会重新发布依赖它的框架包。
- 测试、文档、演示和 CI 从不触发发布。`demos/docx/` 或 `e2e/docx/` 下的更改仅会在 CI 中运行本查看器的检查。

编辑器自身的代码位于 `src/ui/src/docx` 中，并随 `ooxml-ui` 一同发布；该目录的更改会发布 `ooxml-ui`，而查看器软件包通过它们的 `ooxml-ui` 版本范围获取该更改。

## 推送前的检查 {#checks-before-pushing}

在仓库根目录运行 `bun run verify`。对于本查看器，在 `viewers/docx` 中对应的本地检查为：

```sh
bun run typecheck && bun run test && bun run test:scripts
bun run build:packages && bun run check:published && bun run pack:smoke
```

如果某个 tarball 导入了内部包或 `ole2`，`check:published` 会失败；`pack:smoke` 会将 tarball 安装到一个干净的使用环境中，并通过每个框架包加载 DOCX 和旧版 `.doc` 文件。
