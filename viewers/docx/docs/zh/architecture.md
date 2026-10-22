# 架构与职责归属

## 依赖方向

```text
published, one per framework (self-contained bundles)
  docx-react-viewer | -vue-viewer | -angular-viewer | -svelte-viewer | -solid-viewer | -vanilla-viewer
  ------------------------------------------------------------------------------------------------
  internal, private, inlined into every bundle:
        bindings (lifecycle and events only)
                         |
                  web-component editor (canvas text measurer)
                         |
        ooxml-core (external, published)
   /docx/layout   /docx/load (format detection, legacy .doc;
   /collab        ole2 codecs inlined by the core)
        ^
   docx-core (external): thin re-export of ooxml-core/docx and /docx/embedded
```

只有 `docx-core` 和六个 `*-viewer` 包会被发布。`web-component` 和 `bindings` 是 `private` 的工作区包：它们是编辑器的共享源码，从不作为 npm 安装目标。`scripts/build-packages.mjs` 会把它们打包进每个框架包，因此压缩包只导入 `docx-core`、`ooxml-core`、ProseMirror 库及其框架对等依赖。如果内部包或 `ole2` 泄漏到压缩包中，`scripts/check-published-refs.mjs` 和打包冒烟测试会使构建失败。

规范的文档模型、DOCX 解析器和序列化器位于`ooxml-core` 的 `docx` 区域；`docx-core` 是一个轻量的已发布入口，仅重新导出它。可编辑模型是 OOXML 的一个刻意保持精简的投影。加载的会话持有原始包和基线模型，因此保存受支持的编辑时可以保留该投影之外的部件。无改动的保存会返回原始字节。

`ooxml-core/docx/load` 会检测容器签名，并选择 DOCX 或旧版 DOC。其旧版读取器会从 Word 97-2003 文件中提取正文段落，使用 ole2 编解码器（由核心内联）处理 CFB 和二进制 Word 结构，并将提取的文本转换为通用模型。其编辑路径只在通过安全检查时才接受纯段落文本。`ooxml-core/docx/layout` 是不依赖 DOM 的分页引擎；编辑器会注入其基于画布的 `TextMeasurer`。实时协作辅助功能（校验、标识、排序规则）来自 `ooxml-core/collab`；`prosemirror-collab` 客户端和权威端、在线状态插件以及装饰仍保留在编辑器（`ooxml-ui/docx`）中。

`<docx-editor>` 元素（源码位于 `ooxml-ui/docx`，由 `web-component` 重新导出）是唯一的编辑界面。它把模型转换为 ProseMirror，应用编辑事务并修复块标识，发出模型变更，并调用已加载会话的保存函数。当前的命令集支持段落编辑、基本的直接文本格式、对齐、简单表格，以及在连续编辑界面上的缩放。页面布局视图是来自 `ooxml-core/docx/layout` 的近似分页视图，并非 Word 的分页。Shadow DOM 限定了工具栏和页面样式的作用域。注册仅在浏览器中进行，并且是幂等的。浏览器测试通过每个绑定来检验同一份约定。

`bindings` 只负责框架集成：挂载/卸载、属性更新、事件转发和命令式引用。新的编辑器功能应放在 Web Component 中；只有框架生命周期或 API 的要求才能成为修改包装层的理由。

## 两个查看器之间的共享

CFB 读写器和 Word 二进制内部实现的规范版本位于 `ole2` 包（`@christophervr/ole2`）中。DOCX 模型、解析器、写入器以及嵌入式 DOCX API（`docx-core/embedded`），其规范版本位于 `ooxml-core/docx`（由 `docx-core` 重新导出）。任何现代 DOCX 实现都不应放入 `ole2`。

共享的 `ole2` 包不依赖 React、Word 或 PowerPoint 界面，只负责 CFB 和旧版 Word 二进制格式。DOCX 属于 OOXML，归 `ooxml-core/docx` 所有。

## 构建与发布边界

共发布七个包（`docx-core` 和六个框架包；参见[发布](/releasing)），并通过自动化流程发布到 npm。共享的 CFB 包不包含任何 DOCX 编解码器。

`docx-core` 仍是每个框架包的真实依赖，而不是被打包进去：应用会直接从中导入 `createDocument` 和模型类型，因此共用同一份副本可确保两侧的 `DocumentModel` 完全一致。ProseMirror 库同样是真实依赖而非打包内容，这样包管理器就可以为编辑器和应用中任何其他 ProseMirror 代码去重出一份副本。`ooxml-core` 也是真实依赖（它带来布局、加载和协作）。其余内部内容都会被打包，因此下游使用方不需要同级源码检出，也不需要未发布的包；ole2 内联在核心中，不是这里任何内容的依赖。

在将其视为与 Word 功能一致之前，请先查阅对等路线图。软件包构建成功和演示路由可用，并不能证明 Word 渲染、分页或往返保真已经对等。
