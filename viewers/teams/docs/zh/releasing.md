# 发布软件包

共有七个无作用域的包以 **OpenTeams** 这一产品名称发布到 npm（`openteams` 本身属于一个无关的项目，因此每个名称都比较长）。每个包都有**各自的版本线**，仅在实际发生变化时才会递增版本，并拥有各自的 git 标签 `<npm-name>@<version>`（例如 `openteams-server@0.4.0`）。

发布需要在**仓库根目录**中运行，而不是在本文件夹中。规划器（`scripts/release-plan.mjs`）、查看器包表（`scripts/viewer-packages.mjs`）、工作流（`.github/workflows/release.yml`）、受信任发布以及一次性设置，均在根目录的 [`docs/releasing.md`](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md) 中说明。本文件夹没有自己的发布脚本，也没有自己的工作流（其 `.github/` 不生效），也没有 `release:plan` 或 `publish:local` 命令。

| 包                         | 目录               | 依赖（注册表）                                            | 打包进                   |
| -------------------------- | ------------------ | --------------------------------------------------------- | ------------------------ |
| `openteams-react-viewer`   | `packages/react`   | `ooxml-core`、`ooxml-ui`、`lit`；对等依赖 `react`         | `packages/web-component` |
| `openteams-vue-viewer`     | `packages/vue`     | `ooxml-core`、`ooxml-ui`、`lit`；对等依赖 `vue`           | `packages/web-component` |
| `openteams-angular-viewer` | `packages/angular` | `ooxml-core`、`ooxml-ui`、`lit`；对等依赖 `@angular/core` | `packages/web-component` |
| `openteams-svelte-viewer`  | `packages/svelte`  | `ooxml-core`、`ooxml-ui`、`lit`；对等依赖 `svelte`        | `packages/web-component` |
| `openteams-solid-viewer`   | `packages/solid`   | `ooxml-core`、`ooxml-ui`、`lit`；对等依赖 `solid-js`      | `packages/web-component` |
| `openteams-vanilla-viewer` | `packages/vanilla` | `ooxml-core`、`ooxml-ui`、`lit`                           | `packages/web-component` |
| `openteams-server`         | `server`           | `ws`、`yjs`、`y-protocols`、`lib0`                        | 无（发布其源码）         |

`teams-viewer`（`packages/web-component`，即 `ooxml-ui/teams` 的重新导出，其中包含 `<teams-app>` 和绑定连接逻辑）以及各演示均为 `private`，从不发布。`scripts/build-packages.mjs` 会把 Web 组件（包括 CSS）内联进每个框架的打包产物（vite 库构建），并通过 rollup-plugin-dts 将其声明文件展平为每个包的 `dist/index.d.ts`，因此压缩包只导入其清单中声明的内容。Svelte 发布的是源码形式的 `dist/Teams.svelte`，旁边附带已打包的 `dist/runtime.js`（即存储和辅助函数，位于 `openteams-svelte-viewer/runtime` 下）；Angular 是一个已打包的 ES 模块，其装饰器在运行时应用（使用它的应用需要 JIT 编译器）。服务器是普通的 ESM，不经构建；其 `openteams-server` 命令对应 `server/bin.mjs`。

在 monorepo 中，各框架的清单将 `ooxml-core` 和 `ooxml-ui` 声明为 `*`；根目录的发布流程会把这些范围改写为与其同时发布的版本对应的插入符（caret）范围。需要同时修改核心和本查看器的更改，应在同一次推送中落地，并一同发布。

## 决定发布的因素 {#what-decides-a-release-here}

根目录的规划器从 `scripts/viewer-packages.mjs` 中 `teams` 条目的 `bundled` 和 `globals` 读取配置：

- **自身文件。** 包目录下任何非测试文件都会发布该包。测试、`CHANGELOG.md`，以及仅修改 `version`、`scripts` 或 `devDependencies` 的清单编辑，都不会触发发布。
- **已打包的内部包。** `packages/web-component` 下任何非测试的更改都会发布每个框架包。
- **共享构建流水线。** `scripts/build-packages.mjs` 和 `tsconfig.release.json` 的更改会发布每个框架包。服务器选择退出（`global: false`）：只有 `server/**` 会发布它。

版本提升级别取自影响该包作用域内已发布文件的提交中最高的 Conventional Commit 级别：`!` 或 `BREAKING CHANGE:` 为主版本，`feat` 为次版本，其他为补丁版本。请使用 `docs(teams): ...`、`fix(server): ...` 等作用域（参见根目录的 `AGENTS.md`）。

## 针对本查看器运行的检查 {#checks-that-run-for-this-viewer}

`bun run check:published`（已发布的清单中不得有 `workspace:`、`file:` 或 `link:` 范围，也不得有私有包；它还会扫描 `.mjs` 文件和服务器目录）以及 `bun run pack:smoke`（安装并导入压缩包）。`bun run test:scripts` 运行该查看器自己的脚本测试（`scripts/*.test.mjs`）。

CI 运行 `bun run pack:smoke --workspace-runtime`：压缩包使用本检出中的 `ooxml-core` 和 `ooxml-ui` 进行安装（先行构建），这正是发布时同时发布的组合。若不加该参数，`pack:smoke` 会安装注册表中的核心和 UI，即外部用户所获得的版本。根目录的 `verify` 仅在使用 `--full` 时运行这些包步骤。

切勿手动运行 `npm publish` 或推送发布标签，也不要手动修改版本或 `CHANGELOG.md` 文件；发布提交会负责这些。新的包名需要按根目录页面中的说明，完成手动引导和受信任发布者设置。
