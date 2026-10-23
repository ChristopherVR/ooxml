# 架构 {#architecture}

## 职责归属 {#ownership}

- `src/core/visio`（`ooxml-core/visio` 入口）：包的接收、XML、模型、继承、缓存几何、ShapeSheet 求值、编辑命令、协作映射（`ooxml-core/visio/collab`）以及结构化诊断。此处不包含任何查看器 UI。不依赖 DOM 的视图约定同样在核心中，位于 `src/core/visio/ui`：`contract.ts`（属性和事件）、`document-history.ts`（有上限的、以源为依据的编辑历史以及 VSDX 副本导出状态；XML/包的变更通过专用 Worker 运行）、`document-text-search.ts`（对规范化后的可见形状文本进行有上限的字面匹配，给出不可变的、按页面划分的结果）、协作选项与在线状态类型，以及其他不依赖 DOM 的编辑和视图方案。
- `src/ui/src/visio`（以 `ooxml-ui/visio` 发布）：浏览器查看器。`controller.ts` 负责文档/视图状态、选择、事件、“最新加载优先”行为以及事务性编辑；`viewer-element.ts` 是共享的 `<visio-viewer>` 界面；`binding.ts` 是仅限客户端的 mount/update/load/fit/exportSvg/destroy 生命周期（重入的较新更新会取代尚未处理的较旧补丁）；`render-svg.ts` 和 `render-text.ts` 负责 SVG 呈现与渲染警告；`export-svg.ts` 通过同一 SVG 渲染器对当前页面进行有上限的静态序列化；`print-snapshot.ts` 生成无副作用的所选页面产物并带有聚合预算（它不会创建打印框架，也不实现打印机策略）；构建模块是 `building-blocks-*.ts`。
- `viewers/visio/src/index.ts`：对 `ooxml-ui/visio` 的重新导出。
- `src/ui/src/bindings`（以 `ooxml-ui/bindings` 发布）：docx、xlsx、visio 和 teams 共用的框架无关绑定工具包（延迟句柄、属性选取、仅转发变更，以及 React 和 Vue 的组件工厂）。`viewers/visio/packages/bindings/src`（私有）包含 Visio 契约（`VIEWER_CONTRACT`）、`mountFrameworkViewer` 的按修订转发以及每个适配器；`packages/{react,vue,angular,svelte,solid,vanilla}` 是发布包。原生的构建模块组合也在其中。没有针对各框架的解析、渲染或几何分叉。

## 边界 {#boundaries}

引擎目前使用缓存的 ShapeSheet 值。这并非全面的公式求值。浏览器字体度量、文本换行、混合段落间距/缩进、从右到左文本和项目符号均以有上限的工作量实现；这些并不能证明与 Visio 的文本完全一致。

文档 XML/HTML 不会被注入 DOM。SVG 节点通过 DOM API 创建。不提供文档上传、遥测或远程文档获取。包和关系检查属于核心。浏览器组件在专用 Worker 中解析，支持取消，并带有 15 秒的父级超时。场景、元数据、文本工作量和已解码栅格各自独立的预算保护着渲染。图像资源在每次渲染内共享，并在被替换、断开连接或销毁时释放。字体加载监听器会触发文本重排，并在断开连接时移除。这些防御措施不构成安全认证。

实验性的、以源为依据的纯文本替换支持有上限的撤销/重做以及明确的 VSDX 副本导出。核心会拒绝与母版关联的形状、富文本、字段、已签名的包以及宏内容。仅有模型的文档为只读。公式缓存不会被重新计算，在原生 Visio 中重新打开的验证仍待完成；这不是对编辑/保存往返的保证。句柄方法、事件和限制见[查看器 API](/zh/api)。

## 安全与限制 {#safety-and-limits}

核心会校验 OPC 关系，拒绝 DTD/实体和可疑的 ZIP 路径，并对归档输入、数量、声明的展开大小、实际流式输出以及 XML 复杂度设定上限。查看器在分配 Blob 之前检查文件大小，并在绘制之前验证外部提供的场景，包括实际的栅格结构、检测到的 MIME 类型和固有尺寸。栅格验证检查的是结构和校验和，而非完整的熵解码；原生解码器的健壮性依然相关。外部关系绝不会被获取，文档字符串一律作为文本节点，不支持的功能会被报告。

浏览器自定义元素使用带取消功能的专用解析 Worker，并在父级一侧设有 15 秒的限制。无头控制器以及不支持 Worker 的环境使用协作式的核心限制。SVG/文本渲染仍在主线程上运行，并分别设有场景、栅格、元数据和文本工作量的限制。该项目尚未通过安全认证，也未达到生产级加固。

## 嵌入式元文件 {#embedded-metafiles}

嵌入的增强型元文件会接受有上限的记录级兼容性诊断。检查器会在任何转换之前拒绝不支持的格式、样式、映射状态和资源放大。它不会执行嵌入的 WMF 数据、字体、图像、脚本或外部引用。一次性的解析器 Worker 仅通过转换器包，在文档预算和由父级持有的硬性期限之内，转换一个范围很窄、已获准的线条/矩形/椭圆及库存对象子集。除非在隔离的宿主中明确提供受信任的转换器，否则直接的核心解析仍然只用于检查。先前检查的全部五个真实 EMF 媒体部件仍不受支持；仅有结构上的准入并不意味着承诺渲染或 Visio 保真度。

无头核心包含一个经过测试的中性矢量清理器和传输验证器。解析器 Worker 使用转换器处理有上限的基本 EMF 子集，该子集由实时渲染、SVG 导出和不可变的打印渲染共用。非 Worker 的解析器回退路径仍不含转换器。生成的纯路径裁剪测试与本地 Skia SVG 后端一致；已安装的 librsvg 后端会忽略嵌套的 clip-path 交集。仍需进行浏览器和原生 Visio 的对比。参见[详细的采用评审](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/docs/research/emf-adoption-review.md)。

## 验证 {#verification}

核心单元测试和夹具测试证明其所断言的解析/模型用例。查看器测试证明生命周期和视图行为。浏览器测试证明实际的交互。要提出保真度方面的主张，仍需要真正由 Visio 创作的夹具、已知的参考渲染以及量化对比。框架运行时测试与包装器编译是分开的。

## 网站 {#website}

静态的多页面 Vite 构建：主页、指南、功能清单、架构以及独立的演示环境。链接为相对链接，资源均为本地资源。没有第三方字体或脚本。着陆页上没有解析器包。

## 发布 {#publication}

八个 npm 包提供无头 Visio API、带有六个原生框架适配器的共享查看器以及 MCP 工具。它们依赖 `ooxml-core` 和 `ooxml-ui`，这两者位于本仓库中并与它们一起发布（见[版本发布](/releasing)）。仓库的 Pages 构建会把文档和演示部署到 `/ooxml/visio/` 下。
