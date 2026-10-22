---
title: OpenXML 符合性
description: ECMA-376 / ISO/IEC 29500 符合性约定，Strict 和 Transitional 包的检测、规范化与往返处理，以及保存时重映射的命名空间族。
---

# Open XML 符合性约定 {#open-xml-conformance-contract}

本项目面向 PresentationML 包所需的 ECMA-376 / ISO/IEC 29500 部分。本页只涵盖 PresentationML；`ooxml-core` 的 Word 和 Excel 区域有各自的文档。

## 完全支持的含义 {#meaning-of-parity}

只有验证了以下所有适用能力，才能认为某项功能得到完整支持。保留未知 XML 很有价值，但不等于理解或编辑该 XML。

覆盖清单 `OPENXML_COVERAGE` 位于 `src/core/pptx/core/openxml/`，按**四个**维度评估每个结构，类型为 `OpenXmlCoverageFacet`：

| 维度        | 要求                                                             |
| ----------- | ---------------------------------------------------------------- |
| `parse`     | 解析每种符合规范的表示形式，不丢失数据，也不发生未报告的回退。   |
| `preserve`  | 在修改后的保存中保留不支持的标记、关系、内容类型、顺序和包部件。 |
| `edit`      | 在带类型的模型中暴露该功能，执行支持的修改时不破坏无关标记。     |
| `serialize` | 输出能通过所选符合性类别验证的 Strict 或 Transitional 标记。     |

每个维度使用 `OpenXmlCoverageLevel` 评级：`native`、`partial`、`passthrough`、`unsupported` 或 `unassessed`。

::: warning 清单不评价渲染
这里有意**不设置 `render` 维度**。清单描述的是包级往返能力，而非像素效果：某个结构的四个维度都可能是 `native`，屏幕显示却仍采用近似实现。视觉保真度单独记录在[已知限制](/zh/guide/limitations)中，了解幻灯片实际显示效果时应阅读该页。
:::

## Strict 与 Transitional {#strict-vs-transitional}

ISO/IEC 29500 为包内标记定义了两种符合性类别：

- **Transitional**（ECMA-376）：几乎所有 PowerPoint 文件使用的形式，标记命名空间位于 `http://schemas.openxmlformats.org/...`。
- **Strict**（ISO/IEC 29500 Strict）：ISO 推荐的子集，Office 2013 及以上版本可通过“Strict Open XML Presentation”生成。标记命名空间位于 `http://purl.oclc.org/ooxml/...`，根 `p:presentation` 元素带有 `conformance="strict"`。

两种类别对*相同元素使用不同的命名空间 URI*。硬编码 Transitional URI 的解析器会把 Strict 文件看作无法识别的标记，因此许多库完全无法处理 Strict 文件。`pptx-viewer` 双向支持这两种类别。

### 加载时的处理 {#what-happens-on-load}

Strict 处理实现在 `src/core/pptx/core/utils/strict-namespace-map.ts`，并接入运行时状态模块：

1. `detectStrictConformance()` 检查解析后的演示文稿根元素上的命名空间声明（`xmlns` / `xmlns:*`）。存在任意 `http://purl.oclc.org/ooxml/...` URI，就将文件标记为 Strict。
2. `normalizeStrictXml()` **原地**重写已解析的树，将命名空间声明、关系 `Type` 属性和扩展 `uri` 属性转换为对应的 Transitional 形式。
3. XML 解析器包在 Proxy 中，因此整个加载流程后续的每次 `parse()` 调用都会透明地规范化结果。其他代码，包括所有元素解析器、主题解析、图表解析等，始终只看到 Transitional URI，无需感知 Strict。
4. 检测出的类别记录在模型的 `data.conformance: 'strict' | 'transitional'` 中。

### 保存时的处理 {#what-happens-on-save}

`save()` 接受符合性选项：

```ts
const bytes = await handler.save(data.slides, {
	conformance: 'preserve', // default: match the loaded file
	// conformance: 'strict',       // force Strict output
	// conformance: 'transitional', // force Transitional output
});
```

- `'preserve'` 是默认值，使用加载时检测到的类别。因此 Strict 文件经过加载、编辑和保存后，仍然输出 Strict，无需额外设置。
- 实际类别为 `'strict'` 时，保存流程最后一步 `convertZipToStrictConformance()` 重新解析归档中每个 `.xml` 和 `.rels` 部件，原地应用 `convertXmlToStrict()`，转换命名空间声明、关系类型和扩展 URI，并按 Strict 模式要求，在 `p:presentation` 根元素上设置 `conformance="strict"`。无法解析的部件，例如扩展名为 `.xml` 的二进制内容，保持不变；转换会逐部件尽力完成。
- 与符合性类别相关的保存常量，包括关系类型和写入器使用的命名空间，会提前根据实际类别选定，因此新创建部件从一开始就是正确形式，无需事后翻译。

## 重映射范围与保持规范形式的内容 {#what-gets-remapped-and-what-stays-canonical}

只有 ISO/IEC 29500-1 定义的**标记语言族**会在两种符合性类别之间重映射。这与 Open XML SDK 打开 Strict 包时使用的权威转换表一致。

| 命名空间族                                                      | 是否重映射       | Strict URI 示例                                                 | Transitional URI 示例                                                       |
| --------------------------------------------------------------- | ---------------- | --------------------------------------------------------------- | --------------------------------------------------------------------------- |
| `presentationml`                                                | 是               | `http://purl.oclc.org/ooxml/presentationml/main`                | `http://schemas.openxmlformats.org/presentationml/2006/main`                |
| `drawingml` （主命名空间、图表、图形、图片等）                  | 是               | `http://purl.oclc.org/ooxml/drawingml/chart`                    | `http://schemas.openxmlformats.org/drawingml/2006/chart`                    |
| `officeDocument` （包括关系类型 URI、数学、参考文献和文档属性） | 是               | `http://purl.oclc.org/ooxml/officeDocument/relationships/image` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image` |
| `spreadsheetml` （嵌入图表工作簿）                              | 是               | `http://purl.oclc.org/ooxml/spreadsheetml/main`                 | `http://schemas.openxmlformats.org/spreadsheetml/2006/main`                 |
| `wordprocessingml` （嵌入文档）                                 | 是               | `http://purl.oclc.org/ooxml/wordprocessingml/main`              | `http://schemas.openxmlformats.org/wordprocessingml/2006/main`              |
| `schemaLibrary`                                                 | 是               | `http://purl.oclc.org/ooxml/schemaLibrary/main`                 | `http://schemas.openxmlformats.org/schemaLibrary/2006/main`                 |
| `descriptions`                                                  | 是（不同主机名） | `http://purl.oclc.org/ooxml/descriptions/base`                  | `http://descriptions.openxmlformats.org/description/base`                   |
| 开放打包约定（`package/*` 内容类型、关系、核心属性、数字签名）  | **否**           | 两种类别中均保持规范形式                                        | 两种类别中均保持规范形式                                                    |
| 标记兼容性（`markup-compatibility/2006`）                       | **否**           | 两种类别中均保持规范形式                                        | 两种类别中均保持规范形式                                                    |

::: warning OPC 和 MCE 独立于符合性类别
Open Packaging Conventions（ISO/IEC 29500-2）以及 Markup Compatibility and Extensibility（ISO/IEC 29500-3）是共享规范，不依赖符合性类别。真实 Office“Strict Open XML”文件即使在部件内部使用 Strict 的 `purl.oclc.org` 命名空间，仍会让 OPC 关系类型和 `mc:` 命名空间保持规范的 `schemas.openxmlformats.org` 形式。重映射它们会产生 Office 和规范都不接受的文件，因此 `pptx-viewer` 在双向转换中都明确保持它们不变。
:::

### 结构推导规则 {#the-structural-derivation-rule}

除了明确列出的已知对应关系表之外，映射还利用了一项规律：需要重映射的族中，Strict 与 Transitional URI 通过确定性规则关联，并非任意查表：

```
Strict:        http://purl.oclc.org/ooxml/<family>/<tail...>
Transitional:  http://schemas.openxmlformats.org/<family>/2006/<tail...>
```

主机名发生替换，并在族路径段后插入 `2006` 版本段。`pptx-viewer` 通过算法推导这些族中任意 URI 的对应关系，因此即使某个仅 Strict 使用的关系类型或 DrawingML 子命名空间未明确列出，也仍会在加载时规范化、保存时转换回去。`descriptions` 是唯一不遵循此规则的重映射族，它有自己的 Transitional 主机名，因此只放在显式映射表中。

## 往返保证 {#round-trip-guarantees}

- **Strict 输入，Strict 输出。** 加载 Strict 文件时记录 `data.conformance === 'strict'`，使用默认 `'preserve'` 保存时，再次输出 Strict 命名空间和 `conformance="strict"`。
- **无损内部规范化。** 规范化只针对命名空间 URI，不改变元素结构、顺序、属性和未知标记，因此 Strict 处理可以与透传保存行为配合。
- **双向显式转换。** `conformance: 'strict'` 和 `conformance: 'transitional'` 可将包转换为另一类别，不改写与类别无关的 OPC 和 MCE 命名空间。
- **依据真实包结构验证。** 映射本身由 `strict-namespace-map.test.ts` 单元测试覆盖，集成往返测试位于 `src/core/pptx/__tests__/integration/strict-conformance-roundtrip.test.ts`。其中的包模拟真实 Office 创建的 Strict 文件结构，包括这些文件保留的规范 OPC 命名空间。

## 符合性检查门槛 {#conformance-gates}

宣称完整支持 PresentationML，需要通过全部以下门槛：

1. 官方 ECMA-376 Strict 和 Transitional 模式验证每个生成包，以及语料库中每个修改后保存的结果。
2. Open Packaging Conventions 检查覆盖部件名称、内容类型、关系、外部目标、压缩和标记兼容性。
3. 兼容性 API 以稳定代码和 XML 位置报告每个仅保留、回退、有损或不支持的结构。
4. `mc:Choice` 根据已验证的功能能力选择，不能仅因识别了命名空间前缀就选中。
5. 真实 PowerPoint 创建的语料测试强制执行修改后序列化，并比较包结构、类型化语义和参考渲染。
6. Strict 到 Transitional，以及 Transitional 到 Strict 的转换，在不改写无关命名空间的前提下都能通过验证。
7. 发布完全支持的声明前，覆盖清单中不能有任何维度仍为 `partial`、`passthrough`、`unsupported` 或 `unassessed`。

## 覆盖族范围 {#scope-families}

覆盖清单必须包含：

- Open Packaging Conventions 和文档属性。
- 演示文稿结构、幻灯片、母版、布局、备注、讲义、批注、标签、节、自定义放映、视图和演示文稿属性。
- DrawingML 的几何、文本、颜色、填充、线条、效果、变换、锁定、媒体、主题和表格词汇。
- 经典图表、扩展图表、图表绘图和嵌入工作簿。
- DiagramML / SmartArt 的数据、布局、颜色、样式和缓存绘图。
- 时间树、构建列表、切换、触发器、声音和媒体计时。
- 图片、SVG、墨迹、内容部件、OLE、ActiveX、VML、三维模型和扩展。
- 标记兼容性和 Microsoft PresentationML 扩展命名空间。
- Strict 和 Transitional 符合性类别。

## 证据要求 {#evidence-policy}

单元测试证明单个映射，合成包测试证明写入器组合，真实文件修改后保存测试证明互操作性，模式验证证明结构符合性。任何一种证据都不能单独作为完全支持的证明。

## 扩展命名空间与模式边界属性 {#extension-namespace-and-schema-edge-attributes}

有些结构看起来像缺口，检查基础模式的实际声明后才能理解：

**切换时长使用扩展命名空间是合法行为，并非遗漏。** `CT_SlideTransition`（S19.3.1.50，Transitional 模式）只声明 `spd`、`advClick` 和 `advTm` 属性；基础 PresentationML 模式中，没有任何以毫秒表示幻灯片切换时长的 `dur` 属性。PowerPoint 自身需要这个值，因此将其写入 Office 2010 扩展命名空间，作为 `p14:dur`。COM 已验证：在 PowerPoint 2016 中，通过 `Presentations.Add`、`Slide.SlideShowTransition.Duration = 2.5` 和 `SaveAs(ppSaveAsOpenXMLPresentation)`，PowerPoint 会将整个 `p:transition` 包在 `mc:AlternateContent` 中，写出 `<mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="2500" .../></mc:Choice><mc:Fallback><p:transition spd="slow" .../></mc:Fallback>`。对于不带前缀的 `dur`，PowerPoint 不只是容忍，而是静默忽略：只含 `dur="2000"` 的包重新打开后使用 PowerPoint 默认的 0.5 秒，已通过 COM 验证。`pptx-viewer` 也写入 `p14:dur`（`src/core/pptx/core/core/runtime/slide-transition-duration-ns.ts`），但使用更简单的幻灯片根元素 `mc:Ignorable="p14"` 声明，而不是把每个切换包进 `mc:AlternateContent`。PowerPoint 接受两种形式，并都遵循时长，已通过 COM 验证。两种形式都不理解的读取器会回退到 `spd` 速度关键字，因此 PowerPoint 和 `pptx-viewer` 都会在输出 `p14:dur` 的同时保留 `spd`。

**`p:animEffect/@filter="image"` 指定了没有实际载荷的滤镜。** `ST_TLAnimateEffectFilter`（19.5.5）枚举 `p:animEffect` 的 `@filter` 可指定的 SMIL 风格滤镜族，其中包括 `image`。但 `CT_TLAnimateEffectBehavior`（19.5.3）没有为该元素提供任何可承载第二个、独立指定图片引用的子元素或属性。`image` 表示基于图片的擦除或遮罩切换，但时间树和关系中没有任何模式位置可指定使用哪张图片，因此包括 PowerPoint 在内，任何符合规范的读取器都无法恢复预期滤镜。`pptx-viewer` 将它与其他没有专门渲染实现的滤镜族一样处理，回退为中性的淡入淡出。

**`p:bldP/p:tmplLst` 是创作时模板，不是播放输入。** `CT_TLTemplateList`（19.5.84）及其 `p:tmpl` 条目（`CT_TLTemplate`，19.5.85），让文本构建为各大纲级别声明默认计时。`src/core/pptx/core/services/animation-timing-templates.ts` 会以类型化形式解析并往返保留它们，但有意不将其用于播放。其语义是：只有某个大纲级别尚无实例化节点时，PowerPoint 才复制模板的 `p:tnLst` 为它初始化计时，也就是用户正在动画窗格中添加新项目符号时。保存文件中实际存在且可见的段落，在其所属级别下，都已在 `p:timing/p:tnLst` 中有明确节点；PowerPoint 保存前会为每个当前使用的级别实例化节点，因此合法保存的文稿中不存在只有模板覆盖的级别。真实 PowerPoint 创建的语料 `anatidae-animation.pptx`（由 `animation-build-templates-surgical-roundtrip.test.ts` 使用）和完整重建往返测试支持这一结论，但没有通过新的 COM `CreateVideo` 捕获重新验证。要构造当前使用级别缺少自身节点的文件，必须手工编辑时间树，而已知 PowerPoint 在关闭 `DisplayAlerts` 时会静默修复格式不正确的包（参见 `scripts/pptx-com-open.ps1`），这样无论观察到怎样的“模板被忽略”结果，都无法作为可靠证据。

## 组合边界重包围与编辑顺序 {#group-re-wrap-and-edit-order}

位置和尺寸以整数像素暴露，每像素对应 9,525 EMU，同时保留精确源 EMU（`xEmu` / `yEmu` / `widthEmu` / `heightEmu`）。未移动的元素，以及任意嵌套深度中未修改的组合，包括其自身位置、`a:chOff` / `a:chExt` 和每个子元素，在保存时都会逐字节输出原始 `a:off` / `a:ext` / `a:chOff` / `a:chExt`，无论文件使用哪种子坐标空间约定。直接调整组合尺寸时，子坐标空间和每个子元素保持逐字节相同，与 PowerPoint 一致，已通过 COM 验证。

移动或调整子元素尺寸时，会保持组合原有子坐标空间，未触碰的同级元素逐字节不变；同时像 PowerPoint 的边界框自动适应一样，让组合自身的 `a:chOff` / `a:chExt` 和 `a:off` / `a:ext` 紧密重新包围新的子元素集合，并向上传播到每个因此发生边界变化的祖先组合。以下情况已通过 COM 验证为逐字节一致：普通移动；旋转子元素（单独旋转不产生影响，与 PowerPoint 一致）；旋转组合（重现固定旋转中心）；嵌套组合传播；未旋转组合在同一次保存中既直接调整自身尺寸，又编辑子元素；旋转组合在同一次保存中既直接调整自身尺寸，又移动和缩放子元素；以及作为组合直接子元素的旋转形状只调整尺寸而不移动，包括一次编辑中同时调整宽高。最后一种情况在 25 / 37 / -40 / 61 / 113 / 155 / 200 / 290 度均经 COM 验证为逐字节一致。要匹配 PowerPoint，必须将考虑旋转的尺寸调整组合为按轴顺序执行的两次修正，先宽后高，每一步都以上一步结果重新锚定，而非一次性旋转。

直接调整旋转组合尺寸且不修改子元素，也达到逐字节一致。COM 真值覆盖 25 / 90 / 180 / -40 度、同时或分别设置 `Shape.Width` / `Height`，以及从左上角或中心执行 `ScaleWidth`，都符合一条规则：尺寸变化后，单个锚点经旋转在屏幕上保持固定；锚点是未移动的边，或中心锚定缩放时的精确中心。保存流程通过相应移动 `a:off` 重现该规则。同一修复和 COM 验证也覆盖直接调整普通非组合旋转形状的尺寸，包括一次编辑同时调整两轴，使用相同顺序修正和相同 8 个角度扫描。

剩下的一处非直角差距是：旋转组合在**同一次保存**中既直接调整自身尺寸，又移动或调整子元素尺寸。COM 扫描确认了两点。首先，先提交组合自身尺寸调整，再执行子元素编辑和紧密重包围，是此 SDK 单一最终状态保存流程唯一能够实现的顺序。反向顺序最多相差 71,000 EMU，约 0.08 英寸，视觉上可见，不是舍入差异，因为 PowerPoint 在子元素变化时立即重新适应组合边界，后执行的编辑总是基于已经重新适应的中间边界组合计算。从一个最终元素树保存，无法恢复用户两次独立操作谁先谁后，因此反向顺序不属于该修复的目标。在可实现的顺序内，8 个角度乘以 3 种编辑组合（子元素移动、缩放、移动加缩放），共 24 种情况中 9 种逐字节一致；其余在 `a:off` 的 x/y 或 `a:ext` 的宽高上，与 COM 真值最多差 2 EMU，即 2/914400 英寸，没有一致的符号或轴方向。这与上面的顺序修正问题不同：将该修复应用到此情况的自身缩放步骤，24 个结果都没有变化，已验证；以相同方式进一步拆分紧密重包围步骤会变差，只剩 6/24 一致；让中间中心在整条链中保持未舍入浮点数则为 5/24，还破坏了一个原本精确的未旋转情况；全程使用单精度三角函数也没有任何变化。多个情况中 `a:ext` 高度发生漂移，而该公式里的高度根本没有旋转项，这表明 PowerPoint 对这一特定组合使用不同内部流程，而不是可以从黑盒输出推导出来的舍入顺序修正。

后续 COM 实验（提交 `dc6692eba`）证明了这一点：固定子元素**最终**位置和尺寸，只改变 `GroupItems(1).Left` / `Top` / `Width` / `Height` 的赋值**顺序**，四个目标值相同、同一 COM 会话、只保存一次，组合自身保存的 `a:ext` 在 200 度时最多变化 589,402 EMU，即 0.64 英寸；25 度时，仅 `cy` 就相差 115,531 EMU（`1029931` 对比 `914400`）。两种顺序各自都可以逐字节复现，并不存在其中一种“更正确”。PowerPoint 在**每次属性赋值后**都会重新适应组合边界，而不是每个逻辑编辑后只做一次。因此这一组合情况没有唯一正确答案：用户输入四个数值的顺序有多少种，就可能有多少个逐字节正确的答案，而最终元素树并没有可供重放的顺序。上面固定某一种顺序真值后，9/24 精确、偏差不超过 2 EMU 的结果，已经接近单一最终状态保存架构所能达到的极限；即使对这一种顺序消除到零，也无法推广到其他同样合法的顺序。完整调查和证明见 `group-tight-rewrap-own-box.ts`，COM 固定数值见 `group-tight-rewrap.test.ts` 的 `grp1st` 扫描和顺序敏感性样例。

## `.ppt` 导出的能力上限 {#ppt-export-ceiling}

`save(slides, { outputFormat: 'ppt' })` 在 OLE2（CFB）容器中写入真实的 MS-PPT/OfficeArt 记录流（`src/core/pptx/core/ppt/writer/`），不是占位实现。它可无损往返保留幻灯片数量、形状几何、文本和文本片段格式、图片、组合、表格（采用 PowerPoint 2003 自身的分组矩形模型）和背景，明文以及 RC4 CryptoAPI 密码保护（`pptPassword`）两种情况均可；仅通过 SDK 从零构建的演示文稿，可在真实 PowerPoint 16.0 中经 COM 打开，且幻灯片数、形状数和文本一致（已通过 COM 验证，`scripts/com-acceptance-ppt.mjs`）。

**已完成并通过 COM 验证：**

- 形状级和文本片段级超链接或点击操作，包括 URL、跳转到指定幻灯片、所有相对跳转（下一张、上一张、第一张、最后一张、结束放映、上次查看的幻灯片）、命名的自定义放映、`mailto:`，以及打开文件或打开演示文稿，均已对照 `ActionSettings(ppMouseClick)` 验证。
- 存在 PNG/JPEG 预览时，嵌入 OLE 对象（`oleEmbeddedData`）会写为真实 Windows“OLE Package”对象，在嵌套 OLE2 存储中包含 `CompObj` / `Ole10Native`，验证结果为 `OLEFormat.ProgID === "Package"` / `msoEmbeddedOLEObject`。导入器也能读回它：`ExOleEmbedContainer` / `ExOleObjStg`（`src/core/pptx/core/ppt/ole-embed-parser.ts`）会解析为可编辑的 `ole` 元素，已对照一个真实的 PowerPoint 创建的 `.ppt` 验证，该文件带有原生 `Excel.Sheet.8` 嵌入对象（`e2e/pptx/fixtures/ole-embed-excel.ppt`，由 `scripts/make-ole-embed-excel-fixture.ps1` 的 `Shapes.AddOLEObject` 创建，实测 `OLEFormat.ProgID` 为 `"Excel.Sheet.8"`），恢复出的字节经本项目自己的 BIFF8 读取器解码后，得到 PowerPoint 写入的精确单元格值。
- 嵌入音频（WAV）写为真实、可播放的 `SoundCollectionContainer` / `SoundDataBlob`（`media-writer.ts`），甚至超过 PowerPoint 16.0 自身的行为：PowerPoint“另存为 PowerPoint 97-2003”只保留 `SoundContainer` 外壳（`Shape.MediaFormat.Length` 为 0，任何位置都没有 `RIFF` 字节），而本写入器输出的音频，在 PowerPoint 自己另存为 `.pptx` 时，会重新导出为与源 WAV 逐字节一致的内容。
- 视频不会嵌入，但 PowerPoint 16.0 也无法将视频嵌入 97-2003（尝试嵌入后保存会退化为静态图片，`Shape.Type` 为 `msoPicture`），因此写入器现有的图片或占位符降级，已经与 PowerPoint 自身的上限一致。真正链接（外部文件路径）的视频是另一项尚未实现的能力，因为输入字节、输出字节的 `save()` API 没有可供链接的目标目录。
- 墨迹、SmartArt、图表和三维模型现在还会额外写入一个 `metroBlob`（`src/core/pptx/core/ppt/writer/metro-blob-package.ts`）：这正是真实 PowerPoint 2007+ 在 97-2003 文件中写入的同一种 [MS-ODRAW] `OfficeArtTertiaryFOPT` 迷你包（opid `0x03A9`），其中保存元素自身的 OOXML（墨迹为 `p:contentPart`；SmartArt、图表和合成的三维模型框架为改名后的 `p:graphicFrame` / `p:E2oFrame`），因此 PowerPoint 2007+ 阅读器会把该形状重新打开为原本的原生、可编辑对象，而真正的 97-2003 时代阅读器仍会看到本写入器栅格化的预览图片或占位矩形。这些 OOXML 取自本处理器对同一批幻灯片的无损 `.pptx` 保存，因此编辑过的元素往返时保留的是当前状态，而非加载时的状态（`PptxHandlerRuntimeSaveLegacyPpt.ts` 的 `resolveMetroBlobs`）。通过在 PowerPoint 16.0 中重新打开从零生成的 `.ppt` 输出完成 COM 验证：墨迹读回为 `Shape.Type` = 23 / `msoInk`，与 PowerPoint 自身的 97-2003 另存为结果完全一致；SmartArt 读回为 `Shape.HasSmartArt` = `True`，节点数同为 11，文本完整；图表读回为 `Shape.HasChart` = `True`，`ChartType` = 5 / `xlPie`，标题和系列值精确一致；三维模型读回为 `Shape.Type` = 30 / `msoModel3D`，是可用的 `Model3D` 对象。对图表和三维模型而言，这超过了 PowerPoint 自身 97-2003 另存为的上限（见后两项）：PowerPoint 自己在 97-2003 保存时会把两者降级为无法操作的内容，而本写入器的 `.ppt` 输出在 PowerPoint 2007+ 中仍可重新打开为实时、可编辑的对象。携带 `metroBlob` 的降级元素，会标记为信息级的 `ppt-native-roundtrip-<type>` 兼容性警告，而不是普通的 `ppt-unsupported-<type>` 警告，因为此时回退图片或占位符只是 97-2003 专属的限制，并非丢失（`degrade-element.ts`）。
- 对真正的 97-2003 时代阅读器来说，三维模型仍会栅格化为普通图片（`Shape.Type` 为 13 / `msoPicture`，没有 `OLEFormat`），与 PowerPoint 16.0 自身 97-2003 另存为的行为完全一致（用 `scripts/measure-model3d-ole-97.ps1` 测得，该脚本在脚本内构建最小的符合规范的二进制 glTF，并通过 `Shapes.Add3DModel` 插入）；PowerPoint 2007+ 阅读器则会把上面的 `metroBlob` 重新打开为原生 `Model3D` 对象。
- 对真正的 97-2003 时代阅读器来说，图表也仍会栅格化为图片。PowerPoint 16.0 在 97-2003 另存为时，则会将现代图表保留为嵌入的 `Excel.Chart.8` OLE 对象（旧版 MS Graph）（用 `scripts/measure-chart-ole-97.ps1` 测得：`Shape.Type` 为 7 / `msoEmbeddedOLEObject`，`OLEFormat.ProgID` 为 `"Excel.Chart.8"`，`Shape.HasChart` 为 `False`）。为这条仅限 97-2003 的路径写入真实的 `Excel.Chart.8` 对象不在计划内：与 `ExOleObjStg`（有 [MS-PPT] / [MS-ODRAW] 文档，上文已实现）不同，旧版 MS Graph 图表的内部二进制布局没有可供对照验证的公开 Microsoft 规范；上面的 `metroBlob` 对任何 PowerPoint 2007+ 阅读器都完全绕开了这个问题，因为它携带的是真实的现代图表部件。
- 导入不支持 CryptoAPI 之前的 Office 95 RC4/XOR 混淆方案。导入保真度受格式早于 DrawingML 这一事实限制：没有可传递的主题字体方案（转换器会根据文稿收集到的第一个字体合成一个名为“Imported PPT”的方案，找不到时回退到 Arial），也没有二进制对应物的效果会被降级。每个降级元素都会标记一条 `save` 范围的 `PptxCompatibilityWarning`。

- 母版会写入演示文稿自身的文本样式、占位符形状（标题、正文、日期、页脚、幻灯片编号）、由主题推导的配色方案，以及 PowerPoint 2007+ 写入的 `RoundTripTheme12Atom` / `RoundTripColorMapping12Atom` / `RoundTripOArtTextStyles12Atom` 记录（`master-roundtrip-writer.ts`），因此重新打开后的 `SlideMaster.TextStyles`、主题颜色和字体与源演示文稿一致。组合成员按 PowerPoint 的写法写入（`fChild`，以及组合子坐标空间中的 `OfficeArtChildAnchor`），超过 109 个 FAT 扇区的复合文件会获得 DIFAT 扇区，因此包含组合、表格或大于约 7 MB 的演示文稿都能在 PowerPoint 中打开（均已通过 COM 验证）。

**GIF 图片：** 首帧会重新编码为 PNG，其 `IDAT` 经 DEFLATE 压缩并做自适应逐行过滤（`utils/png-encoder-compressed.ts`、`utils/deflate-encode.ts`：LZ77 加单个固定 Huffman 块，同步且可在浏览器中运行），取代了此前的存储块 PNG。单元测试断言结果比存储块编码小得多，且解码后像素完全一致；它尚未通过在 PowerPoint 中重新打开进行 COM 验证，也与 PowerPoint 自己的 zlib 输出不逐字节相同。

**多个幻灯片母版：** 幻灯片遵循不同幻灯片母版的演示文稿，会按首次使用顺序，为幻灯片用到的每个母版各写出一个 `.ppt` 主母版；每张幻灯片的 `SlideAtom.masterIdRef` 按 `[MS-PPT]` 规定，以该母版的幻灯片 ID（`0x80000000` 加其索引）指向自己的母版（`master-roundtrip-writer.ts` 的 `buildDeckMasterRoundTripsFromPptx`）。尚未通过 COM 验证：第一个母版的 ID 是 PowerPoint 自身的单母版值，但后续母版的连续 ID 尚未通过在 PowerPoint 中重新打开来核对。

**仍然存在的 `.ppt` 差距：** 二进制记录写入器本身的限制记录在 [`@christophervr/ole2`](https://github.com/ChristopherVR/ole2#known-limitations-of-the-ppt-writer) 中，由它负责。在本项目一侧：视频以及 WAV 之外的任何嵌入音频格式会降级为图片（见上文）；加密 `.ppt` 的导入仅支持 RC4 CryptoAPI，不支持 CryptoAPI 之前的 Office 95 方案（同样见上文）。

**`metroBlob` 是如何被找到的。** 与图表、视频和三维模型不同，PowerPoint 16.0 在同样的 97-2003 往返过程中仍原生保留墨迹和 SmartArt，因此写入器早先仅把它们降级为图片的做法，并未达到 PowerPoint 自身的上限。墨迹对照一个带有真实 `p14:` 墨迹内容、由 PowerPoint 创建的样例测量（`e2e/pptx/fixtures/ink-contentpart.pptx`，`scripts/measure-ink-ole-97.ps1`）：每个墨迹形状读回时都保持为 `Shape.Type` = 23 / `msoInk`。SmartArt 对照一个由 COM 创建的真实样例测量（`src/core/pptx/__tests__/fixtures/corpus/smartart-orgchart-many.pptx`，`src/core/scripts/pptx/measure-smartart-ole-97.ps1`）：`Shape.HasSmartArt` 在前后均为 `True`（`Shape.Type` = 24 / `msoDiagram`）。

2026-09-11 的从零复现尝试，通过项目自身的 `ole2-parser-read.ts` 和 `record-stream.ts` 读取器分析两个已保存文件，发现 PowerPoint 将两者都表示为普通 MSOSPT 75（“图片框”）形状，其 `OfficeArtTertiaryFOPT` 恰好带有一个复杂属性，未公开的 id `0x3A9`（原始字节 `A9 C3`，即设置了 `fComplex` 和 `fBlipId`），其内容是一个原始 ZIP/OPC“迷你包”：包含真实的 `[Content_Types].xml` / `_rels` 结构，以及未公开但真实存在的内容类型（`application/vnd.ms-office.DrsInk+xml`、`application/inkml+xml`、`application/vnd.ms-office.DrsE2oDoc+xml`、`application/vnd.ms-office.DrsDownRev+xml`）。对墨迹而言，包内的 `drs/inkxml.xml` 是与源文件自己的 `p14:contentPart` 逐字节一致的 `p:contentPart`，另有逐字节一致的 `drs/ink/ink1.xml`。对 SmartArt 而言，`drs/e2oDoc.xml` 是一个 `p:E2oFrame`（改名后的 `p:graphicFrame` / `dgm:relIds`），外加本核心已能无损往返的全部五个图示部件，完全自包含。（这纠正了此前的一个假设：文档级的 `RoundTripCustomTableStyles12Atom`（`0x428C`）与此有关；实际上它只保存一份与这两项功能都无关的通用 `tableStyles.xml` 往返数据。墨迹和 SmartArt 使用相同的 `0x3A9` 属性 id。）

当时，将相同属性（本项目自己生成的包的逐字节副本，其中一次测试还原样拼入了 PowerPoint 自己捕获的字节）写入从零创建的 `.ppt` 中的 MSOSPT 75 形状，**无法**重现 `Shape.Type` = `msoInk` / `msoDiagram`：COM 重新打开后读回的是 `Shape.Type` = 13 / `msoPicture`（带 `pib` blip 引用）或 1 / `msoAutoShape`（不带）。对墨迹做了穷举测试，独立改变精确的 FOPT 属性表、精确的 `ClientAnchor` 位置（取自源文件自己的 `p14:xfrm`），以及形状在 1 至 3 个同级形状中的序号，始终没有得到 `msoInk`。当时的结论是，仅有迷你包并不足够，至少还需要一个尚未识别的信号。

那个“不可能”的结论其实忽略了一件比缺失信号简单得多的事：记录**顺序**。PowerPoint 自身的形状容器把 `OfficeArtTertiaryFOPT` 写在 `FOPT` 和 `ClientAnchor` 之间；此前每次尝试都遵循 [MS-PPT] 对 `SpContainer` 的非正式排序，把 `ClientAnchor` 紧接在 `FOPT` 之后，再追加 `TertiaryFOPT`。在 `buildShapeContainer` 和 `buildPictureContainer`（`shape-writer.ts`）中把它移到二者之间，才是缺失的一环：PowerPoint 随即接受迷你包，并把该形状重新打开为它所描述的原生对象（提交 `b370d429d`、`e1e43448a`）。同一次调查还发现了两个相邻的缺陷，否则真实重新打开后会立刻遇到：`ClientAnchor` 的 16 字节形式的字段顺序是 left / top / right / bottom，而不是此前对 [MS-PPT] 的 `RectStruct` 理解的 top / left / right / bottom（按旧方式写入的探测形状重新打开后会发生转置；`anchor-writer.ts` 现在在所有边都放得下时，也优先使用 PowerPoint 自己始终写入的 8 字节 `SmallRectStruct` 形式），以及绘图的背景形状必须是 `OfficeArtDgContainer` 自己末尾的同级形状，带有 `fBackground | fHaveSpt`，而不是顶层形状组的成员（`drawing-writer.ts`、`sp-container.ts`）。图表和三维模型随后也接入同一套 `metroBlob` 机制（`metro-blob-xml.ts`、`metro-blob-source.ts`、`metro-blob-model3d.ts`），因为现代图表或三维模型同样没有自己的二进制 `.ppt` 表示；重新打开的数值见上文“已完成并通过 COM 验证”。

这次调查还产生了一项无关的、已通过 COM 验证的修复：`wzName` OfficeArt 复杂属性（形状的 `name`）缺少结尾的 UTF-16 空字符，导致真实 PowerPoint 直接拒绝本写入器为具名形状生成的任何 `.ppt`（“Office 检测到此文件存在问题”，且没有修复选项）。已在 `fopt-writer.ts` 的 `encodeComplexString` 中修复，并附带字节级回归测试。

## SmartArt 布局实测基准 {#smartart-layout-ground-truth}

`.pptx` 包含 PowerPoint 自身预计算的绘图部件时，会使用该精确布局，并像 PowerPoint 一样按原始偏移放置，已通过实时 COM 验证。否则由 DiagramML 解释器重建（全部十种 `dgm:alg` 类型、含 `dgm:choose` 门控条目的 `constrLst` / `ruleLst`、相对约束、`presLayoutVars`）。

解释器依据包含 229 个样例的图库测量，这些样例覆盖 `Application.SmartArtLayouts` 报告的全部 176 种内置布局，均由 PowerPoint 自身通过 COM 创建（`src/core/pptx/__tests__/integration/smartart-gallery-ground-truth.test.ts`，通过 `SMARTART_GALLERY_GATE=1` 选择性启用，在通过之前默认跳过）。该门槛把解释器自己的输出（`computeSmartArtElementsWithoutCache`，从不读取缓存）与样例缓存的 `dsp:drawing` 逐形状比较，并按节点文本匹配。

`computeSmartArtElementsWithoutCache` 实际上在这个函数背后运行两个相互独立的引擎：旧版按族划分的解释器（`src/core/diagram/layout/smartart-layout-interpreter.ts`，整个图示选用一种排列器），以及逐数据点的 DiagramML 引擎（`src/core/diagram/engine/`，对每个数据点执行每个 layoutDef `layoutNode` 自己的算法和约束）。`src/core/scripts/pptx/measure-smartart-engine-vs-legacy.ts` 把两者对每个样例独立运行，找出引擎在每个数据集上都严格更准确且不丢失形状集合的布局；`src/core/diagram/engine/engine-first-allowlist.ts` 列出该测量发现的 `layoutDefinition.uniqueId`（例如“Gear”：旧版偏差 76%，引擎 0.1%，因为旧版复合排列器根本无法处理 Gear 中旋转的、装饰性的具名槽位子节点），`computeDiagramMlElements` 仅对这些布局先尝试引擎，其余全部先用旧版，任一方放弃时再回退到另一方。

使用 `src/core/scripts/pptx/gen-smartart-gallery-baseline.ts` 测量（它把逐样例数字写入 `smartart-gallery/baseline.json`；该文件是生成的，不提交）：下面是第十轮所基于的 2026-09-24 运行结果、第十轮的结果，以及第十一轮（2026-09-25，稍后）的结果：

| 检查项                                     | 2026-09-24      | 第十轮          | 第十一轮        |
| ------------------------------------------ | --------------- | --------------- | --------------- |
| 带文本形状集合与 PowerPoint 完全一致       | 229 个中 228 个 | 229 个中 228 个 | 229 个中 228 个 |
| 且所有形状的 x/y/宽/高偏差在图示边界 1% 内 | 229 个中 97 个  | 229 个中 126 个 | 229 个中 143 个 |
| 偏差在 5% 内                               | 229 个中 128 个 | 229 个中 161 个 | 229 个中 181 个 |
| 偏差在 10% 内                              | 229 个中 137 个 | 229 个中 171 个 | 229 个中 191 个 |
| 偏差在 50% 内                              | 229 个中 194 个 | 229 个中 217 个 | 229 个中 226 个 |
| 所有匹配形状的预设形状一致                 | 229 个中 212 个 | 229 个中 222 个 | 229 个中 224 个 |
| 所有匹配形状的字号一致                     | 229 个中 24 个  | 229 个中 114 个 | 229 个中 129 个 |
| 完整门槛（以上全部，按 1%）                | 229 个中 11 个  | 229 个中 97 个  | 229 个中 113 个 |

按图库类别，现在 1% 内的占比为：时间线 7/7、矩阵 3/3、团队介绍 5/6、棱锥 7/10、文本卡片 9/13、列表 28/40、循环 11/17、流程 31/48、层次结构 14/23、关系 19/35、图片 9/27。仍然超出 50% 的三个样例是 Arrow Ribbon、Balance 和 Varying Width List。

三处结构性缺失中有一处早于本轮：`bubble-picture-list--hier5`（特定于预设的空白分隔段落）；`segmented-process--hier5`（`linDir="fromB"` 的堆叠，其条目按 `revPos` 在两种复合模板之间切换）同样是原有缺失，但下面的第九轮把它移到了 `engine-first-allowlist.ts`，完全匹配（偏差为 0），从而关闭了它。另外两处 `numbered-title-list--hier5` 和 `numbered-card-list--hier5` 是下面第八轮的副作用：修复它们真实的形状丢失缺陷后，两者被移到 `engine-first-allowlist.ts`，但引擎仍会用自己（多余的、未折叠的）条目自身文本，为每个强调条目的装饰性背景形状再渲染一次，同时在实际文本叠加层上渲染正确的折叠后自身加后代文本；每个样例因此多出两个无害但不精确的形状，已作为第八轮自己段落中的未解决差距记录在案，而不是不留文档。按图库类别，1% 内的占比为：矩阵 3/3、棱锥 7/10、循环 10/17、层次结构 11/23、关系 15/35、列表 15/40、流程 19/48、图片 7/27，时间线（7）、团队介绍（6）和文本卡片（13）为 0（这是第九轮之前的类别占比；实际变化见下面的第九轮）。

针对 `engine-first-allowlist.ts` 的两轮工作产生了这些数字。首先，引入允许列表（55 个布局）使 1% 内几何匹配从 39/229 提高到 82/229（5% 内 59 到 105，10% 内 72 到 117，50% 内 155 到 178，预设 190 到 204，字号 24 到 25），0 回归（`gen-smartart-gallery-baseline.ts --compare`）。其次，修复 `engine-to-result.ts` 对 `hideGeom` 的处理（ECMA-376 21.4.7.16：`hideGeom` 表示没有可见边框或填充，不是“不是节点”；引擎此前直接丢弃所有此类节点，丢失了按条目角色拆分的形状，例如子节点自身的文本在同级卡片下渲染为第二个无边框方框），恢复了 `vertical-action-list--hier5`（修复前匹配 1/5 个形状，修复后 5/5）和 `descending-block-list--hier5`（1/5 到 5/5），把两者移到允许列表（该修复带来 57 个布局；重新测量其余语料后又发现 `opposing-ideas`、`theme-picture-accent` 和 `theme-picture-alternating-accent` 新增符合条件，共 60 个），使 1% 内几何匹配再提高 82 到 87，同样 0 回归，完整门槛从 10 变为 11。`numbered-title-list` / `numbered-card-list` 遇到的是该修复没有触及的另一个更深层的缺陷（它们的 layoutDef 把子节点文本折叠进父节点的呈现节点，而不是给子节点自己的方框，导致 6 个形状丢失 3 个），仍然优先使用旧版；这是已记录的未解决差距，不是悄悄的变通。完整门槛仍主要被与这两项修复无关的预设或字号不匹配所阻塞。

第三，引擎新增了 `snake` 算法（`src/core/diagram/engine/alg-snake.ts`，ECMA-376 21.4.2.x：子节点填充由节点自身纵横比选定的网格，或固定的 `bkpt="fixed"` 行长，遵循 `flowDir` / `contDir` / `grDir` / `off`，与旧版 `arrangeSnake` 的网格计算一致，但让每个单元格自己的子树通过引擎的常规约束与字号适配流程自行确定尺寸）。据此重新测量，又有 7 个布局符合进入允许列表的条件（共 67 个），同样 0 回归；它们本身都没有越过 1% / 5% / 10% / 50% 的任一档（仍是不够准确的图片或复合族），但其中的 Icon Circle Label List 从旧版偏差 4971% 降到了引擎偏差 64%。

第四，引擎新增了 `cycle` 算法（`src/core/diagram/engine/alg-cycle.ts`，ECMA-376 21.4.2.x：`stAng` / `spanAng` 让子节点均匀分布在环上，`ctrShpMap="fNode"` 把第一个内容点从环上拉到居中的中心，`rotPath="alongPath"` 让每个环上条目旋转到路径自身的切线方向；环半径采用与旧版 `computeCycleRingLayout` 相同的方式求解相邻条目的弦长：先在“自然单位空间”中计算，再做一次各向同性的整体适配，但把环自己的 `sibSp` 间距读作所声明约束的 `fact`，而不是通用约束流程已经解出的值：真实图库 layoutDef 的 `sibSp` 通常把 `refFor="ch" refType="w"` 解析为环上条目自己的“期望”宽度，而同一列表中靠前的 `for="ch"` 约束已将其设为整个图示的宽度，因此解出的值是 `fact * diagramWidth`，而不是 `fact`）。针对 17 个循环或放射样例重新测量，没有新的布局符合进入允许列表的条件：已列出的三个（Cycle Matrix、Segmented Cycle、Radial Picture List）不受影响（偏差相同，因为它们的数据集不会走新的 `cycle` 代码路径），Basic Cycle 和 Multidirectional Cycle 现在与旧版解释器完全持平，而不是回退到旧版（持平不是改进，所以未加入），Block / Continuous / Nondirectional Cycle 和 Radial Cycle 在这个仅有环的简化模型下可测量地变差（没有中心与卫星的间距比、`sibTrans` 凸起或不对称的轴向对齐居中），而 Text Cycle、Hexagon Radial、Basic / Diverging Radial、Radial Cluster、Converging Radial 和 Radial List 仍然直接放弃（复合环上条目，或本轮未建模的中心构造）。这些都不改变门槛数字：0 回归（`gen-smartart-gallery-baseline.ts --compare`），因为这些布局在前后都不是引擎优先。`pyra`、`hierRoot` 和 `hierChild` 在逐点引擎中仍未实现（`registry.ts` 仍用 `composite` 替代它们，由 `isFullySupported` 捕获，因此引擎会干净地放弃，而不是悄悄输出错误的几何）：每个都是实质性的、截然不同的算法族（旧版解释器自己的 `pyramid` / `hierarchy` 排列器加起来仍有数千行），需要各自从零做逐点移植，本轮未尝试。

第五，在尝试 `pyra` 之前，发现并修复了 `constraint-eval.ts` 中的一个单位换算缺陷：真实的 “Basic Pyramid” / “Inverted Pyramid” layoutDef 为其 `level`（以及 `acctBkgd` / `acctTx`）复合子节点声明了一对不带单位的 `ctrX` / `w`（以及 `ctrY` / `h`）字面量：`<dgm:constr type="w" for="ch" forName="level" val="1"/>` 旁边是 `<dgm:constr type="ctrX" for="ch" forName="level" val="1"/>`；引擎通用的“裸字面量即毫米长度”规则把它折叠成约 2.83pt 的方框（实测：`basic-pyramid--flat3` 上引擎偏差 3.212，旧版 0.0019）。对 `val="1"` 的任何单一绝对长度解读都无法得到“填满方框”（在任何统一换算下，`ctrX = w` 都会把它放在 `w/2` 到 `3w/2`），因此，经 COM 验证，这种特定的 `ctrX` / `w`（或 `ctrY` / `h`）裸字面量配对，现在改为按声明节点自身已解析尺寸的分数来读取（`constraint-fill-idiom.ts`），与旧版族解释器独立得出的 `resolveRatioConstraint` 约定一致。在整个 229 个样例的语料中扫描这种精确配对，只在两个棱锥布局中出现，因此该修复在构造上仅限于它们；此前更宽泛的尝试（把每个裸 `w` / `h` / `ctrX` / `ctrY` 等字面量都这样处理）使 Gear（0.0882 到放弃）、Segmented Cycle（0.0012 到 0.3959）、Varying Width List（0.654 到 0.7982）和 Vertical Chevron List（0.3996 到放弃）出现回归，因为这些布局把看起来相同的裸 `val="1"` 用于真正微小的锚点标记（例如 Gear 的 `gear1srcNode` / `gear1dstNode`，即 `hideGeom="1"` 的连接线端点），而不是这种惯用写法；收窄为仅限配对字面量的修复，在全部 229 个样例上精确复现了修改前的测量结果（`measure-smartart-engine-vs-legacy.ts` 和 `gen-smartart-gallery-baseline.ts --compare` 与上一次运行逐字节一致：仍是 1% 内几何 87/229，完整门槛 11/229），因为 `pyra` 本身尚未进入引擎的 `SUPPORTED_ALGS`，所以棱锥布局仍被直接放弃；在实现 `pyra` 之前，这个修复没有可测量的效果，实现之后，它决定了色带是填满自己的槽位，还是折叠成一个点。

第六，引擎新增了 `pyra` 算法（`src/core/diagram/engine/alg-pyra.ts`，ECMA-376 第 1 部分 21.4.2.x：把数据点堆叠为等高的水平色带，每条色带的宽度与其在堆叠中的位置成比例，`linDir="fromT"` 让“Inverted Pyramid”的顶点朝下），移植了旧版 `arrangePyramid` 的色带堆叠几何。真实色带渲染出来后，暴露了第二个单位缺陷：真实的 “Basic Pyramid” 还声明了 “level” 色带自己作用域内的 `<dgm:constr type="h" val="500"/>` / `<dgm:constr type="w" val="1"/>`，而 `preferred-size.ts` 的 `preferredSize`（为保留真实的自声明纵横比而有意如此）会在其父节点已分配的尺寸之上，再次应用复合子节点自己作用域内的尺寸约束，从而覆盖了第五轮的修复，并让 “level” 再次折叠（实测：引擎偏差 1.6079，每个数据集仍放弃）。没有通用规则能把 “level” 的声明与 `radial-cycle` / `segmented-cycle` / `vertical-chevron-list` 自己的 `dummyConnPt` / `wedge*` 节点区分开来，后者声明了相同的裸自身 `w val="1"` / `h val="1"`，且必须保留为真实的约 2.83pt 标记（它们自己的位置来自同级的 `for="ch"` 声明，与 “level” 完全一样），因此 `alg-pyra.ts` 只清理其自身 `pyraLvlNode` 参数所指向的那个具名节点（默认为 `"level"`），在其余树进行布局之前，去掉它裸字面量的自身 `w` / `h`。该修复限定在 `pyra` 算法自己的逐条目子树内，触及不到 `dummyConnPt` / `wedge*`（它们属于不同算法、完全不同的 layoutDef）。有了这两项修复，Basic Pyramid 和 Inverted Pyramid 的 `flat3`（无强调行）数据集现在与旧版完全一致（0.0019，持平）；`hier5` / `hier8`（有强调行）仍为 0.5663，因为 `pyraAcctRatio` / `pyraAcctPos` 强调列拆分（旧版解释器中的 `repositionPyramidBands`）是另一个独立的、按条目角色处理的、复合之后的几何阶段，本轮没有移植；这是已记录的限制，不是悄悄的回归。由于严格的允许列表规则要求布局的每个数据集都改进（或持平）且不丢失形状集合，而 `hier5` / `hier8` 可测量地比旧版更差（0.5663 对 0.0019），所以 Basic Pyramid 和 Inverted Pyramid 本轮都不会新进入 `engine-first-allowlist.ts`。重新测量完整的 229 个样例语料，确认 `pyra` 算法的改动只触及这两个布局（`measure-smartart-engine-vs-legacy.ts` 除这两行外 diff 完全一致），零回归，门槛数字不变（仍是 1% 内几何 87/229，完整门槛 11/229；`pyra` 尚未列入允许列表，所以这两个布局在生产中仍是旧版优先，行为也未改变）。

第七，第六轮留下的 `pyraAcctRatio` / 强调列差距被关闭了：`alg-pyra.ts` 新增了一个后处理阶段（`applyPyraAccentSplit`，由 `layout-driver.ts` 新增的 `LayoutRegistry.resolvePost` 钩子在 `pyra` 节点的子节点完成各自子树布局之后运行），移植了旧版解释器已通过 COM 验证的 `repositionPyramidBands`。一旦 `pyraAcctRatio` 解析为正值（真实的 “Basic Pyramid” / “Inverted Pyramid”，由 `func="maxDepth" op="gte" val="2"` 的 `dgm:choose` 门控），每一行自己的 “level” 色带会缩小到其自然宽度的 `(1 - pyraAcctRatio)`，锚定在图示方框自己的左边缘，而数据点有子节点的行，会得到一个 “acctBkgd” / “acctTx” 列，填满到图示的右边缘。第一次尝试从 “level” 自己缩小后的（宽边）右角开始填充余下部分，在 `hier5` / `hier8` 上测得 0.1142（比第六轮的 0.5663 好，但仍比旧版的 0.0019 差）；把逐形状几何与缓存的基准比较后发现，强调列真正的左边缘是该行缩放后的窄边（顶部）梯形角点，而不是 `arrangePyra` 自己的紧凑边界框所折叠成的宽边角点，因此 `applyPyraAccentSplit` 独立重新计算该角点（与旧版解释器自己的 `pyramidRowGeometry` 一致）。经此修正，Basic Pyramid 和 Inverted Pyramid 的 `hier5` / `hier8` 数据集与旧版完全持平（0.0019，与 `flat3` 相同），因此两个布局都符合“持平或更好”规则，进入 `engine-first-allowlist.ts`，在生产中改为引擎优先。重新测量完整的 229 个样例语料，确认其他地方零回归（`measure-smartart-engine-vs-legacy.ts` 和 `gen-smartart-gallery-baseline.ts --compare` 在两个棱锥行之外不变），门槛数字保持 1% 内几何 87/229、完整门槛 11/229（两个棱锥布局在这一轮之前通过旧版就已落在 1% 内，所以仅路由变化不会改变这些计数）。

第八，`numbered-title-list--hier5.pptx` / `numbered-card-list--hier5.pptx`（真实的 “Numbered Title List” / “Numbered Card List”）丢失了缓存的 6 个形状中的 3 个：“1” / “2” / “3” 序号徽章渲染为空白。它们 layoutDef 的 `sibTransNodeCircle` 通过 `dgm:presOf axis="self" ptType="sibTrans"` 呈现徽章文本，也就是说连接线文本位于**同级**转换点上，而不是**父级**转换点（`smartart-node-tree-axis.ts` 自己的文档注释已为旧版解释器记录了这种不对称：“`numbered-dots-horizontal--hier5.pptx` 的 “1” / “2” / “3” 序号徽章位于 `sibTrans` 点上；与之配对的 `parTrans` 始终为空白”）。`data-points.ts` 的 `buildDataModel` 把连接已解析的 `label` 复制到它合成的 `parTrans` `DataPoint` 上，却从未复制到配对的 `sibTrans` 上，而 `engine-to-result.ts` 只读取被呈现点的内容节点 `.text`（通过 `.source`），从不读取转换点自己的 `.label`，所以从同级一侧呈现的徽章，在这条链的两端都找不到文本。两处都已修复：`sibTrans` 现在像 `parTrans` 一样携带 `label`，并且当没有呈现任何内容点时，`buildRenderedNode` 会回退到被呈现点的 `.label`（以 `RenderedNodeIdentity.literalText` 暴露，与绘图桥接对转换点角色已优先使用的字段相同）。徽章不再空白后，两个布局现在都匹配全部 6 个缓存形状（此前为 3 个），并且在各自唯一的数据集上都可测量地优于旧版（Numbered Title List 0.5854 到 0.3415，Numbered Card List 0.4465 到 0.2458），符合进入 `engine-first-allowlist.ts` 的条件。这是通用的逐点引擎修复，并不仅限于编号列表，因此（按允许列表自己的文档注释，在任何引擎改动之后都要重新测量）重新测量完整语料后，又发现三个无关的布局因同一修复而新符合条件：Converging Arrows 和 Diverging Arrows（旧版 0.0638 到引擎 0.0375，对立箭头对中间类似 “VS” 的 sibTrans 标签）以及 Tabbed Arc（旧版 0.2364 到引擎 0.0131），三者都匹配每个缓存形状且没有多余形状，不同于两个编号列表布局。这两个编号列表样例正是本轮四处（此前为两处）结构性缺失的来源：每个强调条目的 `bgRect`（一个真实的、并未标明如此的背景形状，`dgm:presOf axis="self"`，无 `hideGeom`）会独立呈现同一个“自身”内容点，而该条目实际的文本叠加层（`nodeText`，`hideGeom="1"`，`dgm:presOf axis="desOrSelf"`）也呈现这个点；因此在没有后代的行上，两者渲染相同文本，并在门槛自己的精确文本去重下无害地合并，但在有后代的行上，`bgRect` 未折叠的仅自身文本（“Node One”）与 `nodeText` 折叠后的自身加后代文本（“Node One\nNode Two has a longer label”）是不同的字符串，二者都会保留为各自独立的、未匹配的形状。要彻底查清根因（背景同级形状自己的 `presOf` 何时应被 `hideGeom` 叠加层的 `presOf` 所压制，以匹配真实 PowerPoint 每个条目只有一段文本的实际渲染），是一个关于呈现节点优先级的、独立且更深的问题，不像上面的徽章修复那样属于形状丢失缺陷；该 layoutDef 在 `nodeText` 上自己的 `moveWith="bgRect"` 属性（已解析到 `LdLayoutNode.moveWith`，但从未贯穿到 `PresNode` / `EngineNode`）很可能是这类修复所依据的 ECMA-376 信号，在此作为未解决的、已记录的差距，没有在本轮尝试。使用 `gen-smartart-gallery-baseline.ts --compare` 重新测量完整的 229 个样例语料，确认 1% 内几何和完整门槛完全保持不变（仍是 87/229、11/229，两者均无回归），而结构性丢失从 227 降到 225（上述两个编号列表的幻影形状样例），5% / 10% / 50% 内各有小幅改善（110 到 113，121 到 122，180 到 181），来自箭头和 Tabbed Arc 的改进。

第九，尝试 `hierRoot` / `hierChild`（ECMA-376 第 1 部分 21.4.2.6 / 21.4.2.7，组织结构图和层次结构族）时，暴露了一个大得多的无关缺陷：`engine-to-result.ts` 的 `isFiniteGeometry` 门槛，只要任一渲染节点的方框面积为零，就会放弃**整个**图示，对确实以这种方式渲染且不带文本的形状没有任何例外（`conn` 算法节点自己的 `arrange` 仅对其支持的二维直线情形产生真实方框，`connRout="bend"` 路由时留下 `{w:0, h:0}`；无子节点叶子的 `hierChild` 延续也有意为 `{w:0, h:0}`，因为没有需要扇出的内容）。由于 `SmartArtLayoutResult.connectors` 在下游本来就始终被丢弃（`smartart-interpreter-drawing-bridge.ts` 自己的文档注释：PowerPoint 会自行重建 `dsp:cxn` 形状），`conn` 算法节点本来就不该进入矩形形状收集器，而零面积且无文本的节点无论如何也不会渲染出任何可见内容；现在两者都被排除（`collectRenderedNodes` 完全跳过 `conn` 算法节点；`isRenderable` 现在也直接拒绝非正面积的方框），而不是让整个结果失败。这个与具体布局族无关的修复，一次性让整个语料中此前完全放弃的 40 个布局重新可用；按允许列表自己的“引擎改动后重新生成”规则重新测量后，发现其中 25 个在每个数据集上都可测量地优于旧版且不丢失形状，包括整个 “Meet the Team” 族（此前旧版偏差为 4979%-9979%，是整个语料中最差的，现在引擎偏差低于 83%）以及 `segmented-process--hier5`（0.714 到 0，关闭了上面提到的两个原有结构性缺失之一）。这 25 个现在都进入 `engine-first-allowlist.ts`（共 99 个布局）。使用 `gen-smartart-gallery-baseline.ts --compare` 重新测量完整的 229 个样例语料，确认门槛本身保持不变（11/229），而 1% 内几何从 87 升至 97，5% 内 113 到 128，10% 内 122 到 137，50% 内 181 到 194，预设 204 到 210，结构性 225 到 226（扣除上述编号列表幻影形状的代价，并计入 segmented-process 的收益）；字号匹配数从 25 降到 24（一个新近改为引擎优先的样例字号略有偏差，而旧版恰好精确；没有进一步追究，因为允许列表自己的纳入规则不要求字号一致，而几何方面的收益是决定性的）。随后真正移植了 `hierRoot` / `hierChild`（`src/core/diagram/engine/alg-hier.ts`）：真实的 layoutDef 会按代手工展开递归（`hierChild1` 把顶层点扇出为 `hierRoot1` 条目，每个条目嵌套一个 `composite` 内容框，以及更深一代的、自己的嵌套 `hierChild2` 延续，依此类推），并且每一代都必须以完全相同的尺寸渲染其条目；对每一代独立计算（取各个 `hierChild` 调用自己的局部方框）会得到正确的拓扑，但缩放比例错误，所以最外层的 `hierChild` 调用改为遍历整棵子树一次（按叶子加权的宽度单位和代数深度），推导出一个共享的条目尺寸，并在放置任何后代之前把它设置到每个后代上。只实现了普通的 “std” 扇出行情形（`hierBranch` 悬挂列、`chMax` / `chPref` 换行，以及 `orgChart` 助理槽位，各自是旧版解释器自己约 8800 行的 `smartart-hierarchy-*.ts` 族中独立的、按条目角色处理的分支，均未尝试）。针对整个层次结构族（8 个布局，14 个样例）测量：相比此前的完全放弃有实质性改进（例如真实的 “Hierarchy” 从 0.7767 降到 0.2687，“Horizontal Hierarchy” 从 1.0 降到 0.4925），但在每个数据集上仍可测量地比旧版解释器自己更成熟的实现更差（旧版在该族上为 0.0019-0.3）；因此本轮**没有**把 `hierRoot` / `hierChild` 加入 `engine-first-allowlist.ts`，这是“已尝试，尚不足够”的诚实结论，与第四轮自己的循环尝试相同。无论哪种情况都是零回归，已由上面同一次完整语料重新测量确认。

第十轮（2026-09-25），四项改动，每项都在整个语料上重新测量，没有样例离开 1% 档（`gen-smartart-gallery-baseline.ts --compare`）：

- **字号。** 引擎此前对每个文本节点单独适配，使用单行 0.6em 的猜测值和 18pt 默认值。现在从每个 `tx` 节点的 `primFontSz` 约束开始，按整数磅缩小，直到它的 `dgm:rule` 下限（没有规则则不缩小），直至文本适合该预设自己扣除边距后的文本矩形；`op="equ"` 组以及共享同一个布局节点名称的所有节点（514 个缓存 presName 组中的 510 个以同一字号渲染）取其中最差的适配结果，而 `primFontSz lte` 链接仅在被引用节点缩小时才会限制某个节点。测量本身对照 COM 中的 PowerPoint 做了修正：Aptos 字宽表此前是对自动调整大小方框的 `Shape.Width` 取差，读数偏宽约 1.1% 且没有字距调整，因此 `scripts/make-font-advance-table.ps1` 现在测量 `TextRange.BoundWidth` 外加每个字对的字距（在 12pt 及以上应用，复现的整段标签 COM 宽度误差在 0.2pt 内）；SmartArt 对字体自己 1.2207em 的行高应用 90% 行距（五个样例上每行为 1.099-1.100 倍字号），而不是文本框的 1.2em，再加每块固定的 0.00695em；边距单位为磅（按长度引用的边距，其长度以毫米计）；未声明的边距一侧默认为 0.56 倍字号；`stBulletLvl` 决定哪些段落是项目符号，它们以 `secFontSz`（0.78 倍）渲染。仅就缓存的几何而言，在该模型下，950 个缓存文本形状中有 941 个以其缓存字号适合（此前为 900 个）。旧版解释器共享这些度量修复，其分层适配现在按整数磅向下遍历。
- **层次结构。** `alg-hier.ts` 现在在定义的未缩放空间中对整棵树测量一次，按轮廓而不是边界框来打包同级子树（悬挂的子节点会伸出父节点之外而不加宽该行，`smartart-orgchart-nested-hang.pptx` 中的团队位于其父节点的叶子同级节点之下），把每个节点放在它的 `hierAlign` 一侧（`fromT` 列上的 `tL` / `tR` 节点会把这些列向内悬挂 `0.25 W`，即组织结构图中明确写出的 `init` / `l` / `r` / `hang` 分支），然后整体缩放一次以适配，并以实际绘制的内容居中。Organization Chart（0.0231 到 0.0019）、Horizontal Organization Chart（0.0225 到 0.0019）和 Horizontal Multi-Level Hierarchy（0.2987 到 0.0012）改为引擎优先，Hierarchy 和 Horizontal Hierarchy 也是如此（0.0019 的几何持平，由每个字号都匹配而打破）。带有助理点的图示仍交给旧版解释器。
- **蛇形网格。** `alg-snake.ts` 现在让每个单元格保持其约束尺寸，用 `sibTrans` 间隔的宽度和节点的 `sp` 作为间距，并选取使网格缩放最大的行长（只缩小，从不放大）：Basic Block List、Bending Process 和 Bending Picture 族、Picture Grid 以及 Text Card 族改进到 0-0.0225。
- **形状。** 没有 `type` 的 `dgm:shape` 不再绘制幻影圆角矩形，而第一个被呈现的点是空占位符的无边框方框仍会绘制（点列表布局）。

`src/core/scripts/pptx/measure-smartart-engine-vs-legacy.ts` 现在还会按字号匹配来打破几何持平（Basic Cycle、Titled Matrix），`engine-first-allowlist.ts` 列出 131 个布局（此前为 99 个）。

确实成立的是：核心的单一算法布局（基本流程、垂直流程、基本块列表、基本/连续/多向循环、基本和发散射线、层次结构、水平层次结构、基本和倒棱锥、组织结构图，以及自第十轮起的上述蛇形网格和文本卡片），在语料的每个数据集上都能在 1% 内复现 PowerPoint 的几何，且 229 个样例中有 129 个匹配每个缓存字号（在下面的第十一轮之后）。文本自动适配遵循实测规则（整数磅字号、真实的框架边距和圆角内缩、折叠子段落固定的 0.78 比例），组织结构图另外由 `smartart-orgchart-genuine-fixture.test.ts` 固定（拓扑、悬挂尾部偏移、扇出与单列的选择）。

第十一轮（2026-09-25，稍后），三项改动，每项都在重新测量后没有样例离开 1% 档（`gen-smartart-gallery-baseline.ts --compare`），并重新生成了允许列表（131 到 143 个布局）：

- **由文本驱动的尺寸。** 几何依赖其文本字号的节点（子节点的 `h` 为 `fact x primFontSz`，或 `dgm:rule type="h" val="INF"` 让 `tx` 节点随文本增长）不再折叠：拥有 `primFontSz` 规则的节点从其约束开始按整数磅向下搜索，直到所有文本都适合其（增长后的）方框且没有溢出，字体相对长度按候选字号解析，增长受该节点自己的 `lte` 约束限制，并在 `op="equ"` 组之间保持一致。增长后的方框等于测得的文本，加上预设文本矩形内的边距，再加 0.04em（缓存的 30、35 和 46pt 单行：0.043、0.040、0.040em）。Vertical Bullet List、Vertical Box List 和 Horizontal Bullet List 改进到 0.0019，Basic Chevron Process 到 0.0075，Sub-Step Process 到 0.0281，均为引擎优先。
- **组织结构图助理。** 助理位于管理者下方一个 `sp` 的色带中，且在其他每一行之上，在管理者中心线两侧成对排列（左侧的先放，距中心线 `sibSp / 2`），更多的助理对每对低一个 `sp`；同样的两列放置方式也用于排布 “Both Hanging” 管理者的下属，悬挂列的缩进读取管理者的 `alignOff`（有助理时为 0.65），每个管理者自己的 `hierBranch` 从其呈现点读取。语料中全部 21 个带助理的组织结构图（1 至 4 个助理，全部四种 `hierBranch` 取值，为此通过 COM 创建了 `smartart-orgchart-assistants.pptx`）都在 0.0017 内且所有字号都匹配（旧版为 0.22-0.47）；Half Circle Organization Chart 改进到 0.0012。
- **规则收缩的网格和行。** 当 `dgm:rule type="w"` 让单元格收缩、而单元格随内容增长时，蛇形布局会按每个候选宽度重新布局：PowerPoint 在规则的 `val` 与受约束宽度之间用五次探测做二分，保留最后一个适合的结果（在 Meet the Team 的 flat3 和 hier5 上 EMU 精确）。其流向引用无法缩短的 lin 行，会同时缩小两个引用轴，而沿流向垂直方向不由任何东西确定大小的嵌套 lin，其厚度就是其内容的厚度。Meet the Team 1% 内从 0/6 提高到 5/6，Icon Circle Label List 从 0.4152 到 0.0012，Process List 从 0.8762 到 0.0012，Vertical Equation 从 0.1745 到 0.0012。

仍未解决：Meet the Team hier8（0.0496：角色文本增长的卡片在探测中没有建模）、姓名行的粗体字重（没有 Aptos Bold 度量，所以 Meet the Team 的姓名字号大了一磅）、Name and Title Organization Chart（0.0415）、带标签的层次结构和表格式层次结构、Circle Picture Hierarchy、带助理的水平组织结构图（仍用旧版），以及三个超出 50% 的样例（Arrow Ribbon、Balance、Varying Width List）。这只影响保存时不带缓存绘图的演示文稿；PowerPoint 自己会写入缓存绘图，所以它适用于其他生成器生成的演示文稿，以及结构编辑后的实时重新布局。

## 相关阅读 {#related-reading}

- [已知限制](/zh/guide/limitations)：当前尚未解决的差距。
- [架构](/zh/guide/architecture)：符合性处理在加载和保存流程中的位置。
