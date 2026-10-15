# 架构与职责划分

## 依赖方向

```text
published, one per framework (self-contained bundles)
  @christophervr/xlsx-react-viewer | xlsx-vue-viewer | xlsx-angular-viewer
  xlsx-svelte-viewer | xlsx-solid-viewer | xlsx-vanilla-viewer
  ----------------------------------------------------------------------------
  internal, private, inlined into every bundle:
        bindings (lifecycle and events only)
                         |
               <xlsx-editor> web component (grid painter, ribbon, dialogs)
                         |
        ooxml-core (external, published)                 ooxml-ui (external)
   /xlsx       model, formula engine, layout, edit,      office-ui-* controls
               read/write .xlsx
   /xlsx/load  format detection, .xls (ole2 codecs inlined), CSV
        ^
   @christophervr/xlsx-core (external): thin re-export of ooxml-core/xlsx and /xlsx/load
```

只发布 `@christophervr/xlsx-core` 和六个框架软件包。`web-component` 和 `bindings` 是 `private` 的工作区软件包：它们是编辑器的共享源码，不是 npm 安装目标。`scripts/build-packages.mjs` 将它们打包进每个框架软件包，因此发布的 tarball 只导入 `@christophervr/xlsx-core`、`ooxml-core`、`ooxml-ui` 及其框架 peer 依赖。如果内部软件包或 `ole2` 泄漏进 tarball，`scripts/check-published-refs.mjs` 和打包冒烟测试会使构建失败。

## 核心负责逻辑

所有非 UI 的内容都位于 `ooxml-core` 的 `xlsx` 区域（源码在 `ChristopherVR/ooxml` 仓库）：

- 工作簿模型（`Workbook`、`Worksheet`、`Cell`，以及已解析并去重的单元格样式），
- 地址与区域（`parseAddress`、`formatRange` 等），
- 数字格式（`formatValue`、`parseCellInput`），
- 公式解析器与计算引擎（`createCalcEngine`、依赖跟踪、动态数组），
- 带撤销和重做的编辑命令（`createEditSession`），
- 不依赖 DOM 的视图逻辑：网格度量、可见单元格、单元格视图、条件格式、图表视图模型，
- `.xlsx` 读取器和写入器，以及位于 `/xlsx/load` 下的格式检测、旧版 `.xls` 读取器（通过共享的 ole2 编解码器）和 CSV。

加载的工作簿会把模型无法表示的部分保存在 `workbook.source` 中，因此保存时会原样带过。

## Web 组件负责绘制

`packages/web-component` 是唯一的编辑界面：它在工作簿之上创建编辑会话，绘制核心布局函数返回的内容，把键盘、鼠标和功能区输入路由到编辑命令，并负责样式、语言区域和对话框。Shadow DOM 限定其样式作用域；注册只在浏览器中进行，且是幂等的。浏览器测试通过每一种绑定验证同一份契约。

`packages/bindings` 只负责框架集成：挂载与卸载、属性更新、事件转发和命令式句柄。新的编辑器功能应放在 Web 组件（或核心）中；只有出于框架生命周期或 API 的需要，才应修改绑定。

## 构建与发布边界

`@christophervr/xlsx-core` 仍然是每个框架软件包的真实依赖，而不是被打包进去：应用直接从它导入 `createWorkbook` 和模型类型，因此共享同一份副本可使两侧的 `Workbook` 保持一致。`ooxml-core` 和 `ooxml-ui` 同样是真实依赖。其余内部内容都被打包；ole2 内联在核心中，不是这里任何东西的依赖。参见[发布](/releasing)。
