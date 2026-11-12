# 发布软件包

Excel 软件包由仓库中唯一的自动化流程发布，而非由它们各自的流程发布。本查看器是 [ooxml](https://github.com/ChristopherVR/ooxml) 仓库中的 `viewers/xlsx`，与 `ooxml-core`（`src/core`）和 `ooxml-ui`（`src/ui`）位于同一仓库。整个过程（发布如何决定、版本升级级别、每小时运行的工作流、受信任发布以及一次性的仓库设置）都在仓库根目录的 [docs/releasing.md](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md) 中有所描述。

## 从此处发布的内容 {#what-is-published-from-here}

每个软件包都有自己的版本线、Git 标签 `<npm-name>@<version>` 和更新日志。它们在仓库根目录的 `scripts/viewer-packages.mjs` 中声明。

| 软件包                             | 目录               | 说明                                                 |
| ---------------------------------- | ------------------ | ---------------------------------------------------- |
| `@christophervr/xlsx-core`         | `packages/core`    | 基于 `ooxml-core` 的工作簿模型入口；本身不包含逻辑。 |
| `@christophervr/xlsx-react-viewer` | `packages/react`   | 自包含的编辑器软件包；打包了 Web 组件和绑定。        |
| `xlsx-vue-viewer`                  | `packages/vue`     | 结构相同。                                           |
| `xlsx-angular-viewer`              | `packages/angular` | 结构相同。                                           |
| `xlsx-svelte-viewer`               | `packages/svelte`  | 结构相同。                                           |
| `xlsx-solid-viewer`                | `packages/solid`   | 结构相同。                                           |
| `xlsx-vanilla-viewer`              | `packages/vanilla` | 结构相同；同时负责普通的 `<xlsx-editor>` 入口。      |
| `xlsx-viewer-mcp`                  | `mcp`              | 独立的 MCP 服务器，参见其 README。                   |

内部软件包 `packages/web-component`（`xlsx-web-component`）和 `packages/bindings`（`xlsx-bindings`）为 `private`，从不发布：`scripts/build-packages.mjs` 会将它们打包进每个框架软件包，因此这两者中任何非测试的改动都会同时发布全部六个软件包。已发布的 tarball 只导入 `@christophervr/xlsx-core`、`ooxml-core`、`ooxml-ui` 及其框架 peer；`ole2` 被内联到 `ooxml-core` 中。`ooxml-core` 和 `ooxml-ui` 是真正的注册表依赖，从不被打包。

## 从此目录运行的检查 {#checks-from-this-folder}

在根目录执行 `bun install` 之后，从 `viewers/xlsx` 运行以下命令：

```sh
bun run build:packages
bun run check:published   # no tarball imports an internal package or ole2
bun run pack:smoke        # install the tarballs into a clean consumer and load .xlsx and .xls
```

发布计划器、发布保护和工作流都从仓库根目录运行（`bun run release:plan`、`node scripts/publish-released.mjs --dry-run`、`.github/workflows/release.yml`）。切勿手动运行 `npm publish` 或推送发布标签。

## 提交约定 {#commit-conventions}

版本升级级别取决于提交类型（`!` 或 `BREAKING CHANGE:` 为主版本，`feat` 为次版本，其他为补丁版本），而提交所涉及的文件决定某个软件包是否发布。规则和示例见 [CONTRIBUTING.md](https://github.com/ChristopherVR/ooxml/blob/main/CONTRIBUTING.md#commit-conventions)。
