# 发布

Visio 的软件包由 `ooxml` 仓库自身的流程发布，该流程在仓库的 `docs/releasing.md` 中有所描述：根目录的 `scripts/release-plan.mjs` 会根据自上次 `<npm-name>@<version>` 标签以来的约定式提交为每个软件包确定版本号，根目录的 `.github/workflows/release.yml` 负责打标签、生成更新日志并通过 npm 受信任发布（OIDC 和来源证明，无需令牌）进行发布。Visio 软件包表即根目录 `scripts/viewer-packages.mjs` 中的 `viewer('visio', ...)`（`viewers/visio/scripts/viewer-packages.mjs` 是供构建脚本使用的视图）。本查看器的 `.github/` 文件夹不起作用。

共发布八个软件包：`visio-core` 以及 `visio-{react,vue,angular,svelte,solid,vanilla}-viewer`，分别来自 `packages/{core,react,vue,angular,svelte,solid,vanilla}`，另有 `visio-viewer-mcp` 来自 `mcp`。对此处 `src/` 或 `packages/bindings/src/` 的更改，或对 `scripts/build-packages.mjs` 或 `tsconfig.build.json` 的更改，会发布全部六个查看器软件包；对某个软件包自身目录的更改，只会发布该软件包。查看器的实现本身是 `ooxml-ui/visio`（`src/ui/src/visio`），随 `ooxml-ui` 一同发布；这些软件包通过工作区范围依赖 `ooxml-core` 和 `ooxml-ui`，发布流程会将每个依赖改写为与其一同发布的版本的插入号范围。文档和纯测试的更改不会发布软件包，除非它们修改了已发布软件包的 README。

在 `viewers/visio` 中运行 `bun run check`，会针对注册表中的 `ooxml-core` 和 `ooxml-ui` 验证源代码、适配器、文档、构建以及打包后的使用者（`bun run test:packages`）。CI 会以 `--workspace-runtime` 运行打包后的使用者（`bun run test:packages:workspace`），将本检出中的 `ooxml-core` 和 `ooxml-ui` 一并打包，与发布时的方式相同；`bun run test:browser` 验证浏览器中的行为，并以同样的打包检查在浏览器中结束（`test-packages-browser.mjs --workspace-runtime`）。构建包含两个浏览器 Worker，生成 TypeScript 声明，并提供原生 Svelte 源代码以及 Solid/Svelte 服务端入口。

切勿手动运行 `npm publish` 或推送发布标签。发布失败后，可通过发布工作流的 `tag` 输入重试，该操作会跳过已在 npm 上的版本。

发布工具的来源记录在 `scripts/RELEASE-PROVENANCE.md` 中。
