# 查看器 API {#viewer-api}

每个查看器包背后的、与框架无关的 API。每个框架包（`visio-react-viewer`、`visio-vue-viewer`、`visio-angular-viewer`、`visio-svelte-viewer`、`visio-solid-viewer`、`visio-vanilla-viewer`）都转发相同的句柄方法和事件；各适配器原生的属性、事件和生命周期请参阅 `packages/bindings/README.md`。渲染和编辑均为测试版功能；[功能清单](/parity)列出了尚缺的部分。

## 挂载与加载 {#mount-and-load}

```ts
import { mountViewer } from 'visio-vanilla-viewer';
const viewer = mountViewer(container, {
	zoom: 1,
	events: { 'document-error': (error) => console.error(error) },
});
await viewer.load(file); // Blob, Uint8Array or ArrayBuffer
viewer.fit();
viewer.update({ pageIndex: 1 });
const snapshot = viewer.exportSvg(); // SVG string, dimensions, byteLength and diagnostics
const prepared = viewer.createPrintSnapshot(); // Frozen current-page artifacts, no printing
viewer.destroy(); // idempotent; later commands reject
```

导入该模块对 SSR 是安全的；挂载和元素注册仅限浏览器。`ViewerController` 不依赖 DOM，因此其他渲染器可以订阅相同的状态和事件。`renderPage` 是 SVG 渲染器端口。模块的职责归属见[架构](/zh/architecture)。

## 属性与事件 {#properties-and-events}

每个绑定都接受相同的属性并报告相同的事件（约定为 `ooxml-ui/visio` 中的 `ViewerProperties` 和 `ViewerEvents`；[绑定](/bindings)列出了各框架的原生形式）。

| 接口 | 约定                                                                                                                                                                                   |
| ---- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 属性 | `document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `readOnly`, `ribbonAddIns`, `collaboration`, `storage`, `customization`                                                     |
| 事件 | `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in`, `collaboration-change`, `file-command` |

- `canEdit`（布尔值）是编辑权限。完整查看器默认可编辑，`canEdit: false` 使其只读；[构建模块](/building-blocks)在宿主传入 `true` 之前是只读的。没有包字节的图纸（作为 `document` 传入的模型、旧版 `.vsd`）无论 `canEdit` 取何值都是只读的。
- `document-change` 携带 `{ document, dirty, kind }`。`kind` 对本窗口自己的更改为 `edit`、`undo` 或 `redo`，应用协作者的更改时为 `remote`，`markClean()` 之后为 `save`。
- `markClean()`（位于 `viewer.controller` 和构建模块上）声明当前内容已保存。自行保存 `exportVsdx()` 结果的宿主在保存后调用它：`edit.dirty` 变为 `false`，触发一次 `kind` 为 `save` 的 `document-change`，之后的编辑和撤销步骤以该点为基准计算。对无法回存的图纸，它不执行任何操作。
- `collaboration`（`VisioCollaborationOptions | null`）在设置期间于房间中共享已打开的图纸。共享开始或结束，以及状态、同步或参与者发生变化时，`collaboration-change` 会携带一个 `CollaborationState`。在元素上（句柄的 `viewer.element`），`collaborationState` 读取当前状态，`reconnectCollaboration()` 和 `stopCollaboration()` 对其起作用。参见[协作](/zh/collaboration)。

## 选择形状 {#select-shapes}

每个已挂载的/原生的句柄都提供 `selectShapes(targets)`、`selectAll()` 和 `clearSelection()`。目标的形式为 `{ id, name, pageId? }`；省略 `pageId` 表示当前页面。`controller.state.selectedShapes` 是有序的不可变数组，其第一项为主选形状 `selectedShape`。现有的 `shape-select` 事件报告该主选项；`selection-change` 报告完整数组。这些状态可通过各框架原生的 hook、composable、signal 或 store 获得。Vue 会发出 `selection-change`；Angular 提供 `selectionChange`；六个适配器都接受 `events['selection-change']`。

按住 Shift、Ctrl 或 Meta 点击会切换某个形状。点击背景和按 Escape 会清除选择。隐藏或已删除的目标会被移除，选择某个组时会排除其后代，切换页面或文档会清除选择。全选包含可见的前景形状。针对单个目标的文本和几何面板要求恰好选中一个形状。“开始”选项卡中的格式设置和删除操作会对多个选中的形状以原子方式执行；任一目标受源保护都会使整批操作被拒绝。

## 创建空白绘图 {#create-a-blank-drawing}

每个已挂载的/原生的句柄和共享元素都提供 `createBlankDrawing(options?: CreateVsdxOptions): Promise<void>`。控制器提供相同的方法。可选的 `width` 和 `height` 以物理英寸为单位；默认值为 8.5 乘 11，包含一个空页面，绘图比例为 1:1。“文件 > 新建 > 空白绘图”和 Ctrl+N 调用相同的生命周期。可编辑控件保留其原生快捷键行为。

被接受的“新建”请求会替换文档、选择和历史记录，元素会将其命名为 `New drawing.vsdx`。源字节可以立即导出；`dirty` 记录之后的编辑。被取代的加载、取消和销毁都无法接受迟到的结果。状态变化通过各绑定现有的响应式状态 API 传递。

## 指针移动与框选 {#pointer-movement-and-marquee-selection}

拖动符合条件的本地 2D 形状会按统一的平移量移动当前选择。四个客户端像素的阈值可保留普通点击。预览只改变 SVG 的呈现；释放时提交一次原子的源编辑。Escape、指针捕获丢失、源/页面/选择变化以及工具切换都会取消它。移动使用物理页面坐标，并在核心中转换为绘图单位，同时保留已保存的旋转和本地销钉。源保护始终具有最终决定权。

拖动页面空白处会选中被完全包围的可见前景形状，依据的是它们变换后的宽/高范围。按住 Shift、Ctrl 或 Meta 会加入当前选择。指针移动目前不包括母版、组、图层、外来对象和已粘合的连接线。调整大小手柄和连接线路由是独立的功能；调整大小手柄在下文中说明，而路由仍在此手势的范围之外。

## 调整大小手柄与“大小和位置” {#resize-handles-and-size-and-position}

共享的指针工具会为一个符合条件的本地 2D 形状显示八个调整大小手柄。边缘手柄改变一个维度；角点手柄同时改变两个维度。对边或对角在形状已保存的旋转/翻转后的本地坐标轴中保持固定。已保存的偏离中心的本地销钉在源证据允许时按比例缩放。拖动期间，只有选择框预览新的边界；释放时提交一次源编辑和一个撤销步骤。越过对侧手柄时，该维度会被限制为其原始尺寸与 1/16 物理英寸中的较小者，且不会改变翻转标志。吸附、纵横比修饰手势、完整内容预览以及多形状调整大小均未实现。

“视图 > 任务窗格 > 大小和位置”会打开共享的数值窗格。X/Y 是旋转销钉在绘图英寸中的位置，以左下角为原点，Y 向上增加。宽度和高度同样使用绘图英寸，与物理页面比例无关。角度以度为单位，逆时针为正。按 Enter 或离开某个字段会提交该字段一次；按 Escape 会恢复其已保存的值。与锚定的指针手柄不同，此窗格中的宽度/高度更改会固定旋转销钉。输入与当前值完全相同的值不会改变源字节、公式和历史记录。草稿在缩放以及同一源/选择上的其他编辑期间会保留；替换源或选择其他对象会清除草稿。

这两种控件都要求以源为依据的 VSDX，以及一个带有已保存销钉的、普通的本地、未分层、未粘合的 2D 形状。组、母版、外来对象和多选均不可用。场景资格是初步的：核心在提交时检查源几何、保护、公式和依赖项。被拒绝时会显示错误，而不改变源字节或历史记录。取消、销毁以及更新的源/页面/选择意图会阻止待处理的编辑被接受。

核心的 `visioSizePositionState(page, shapeId)` 辅助函数提供冻结的数值状态。`visioSizePositionCommand(page, shapeId, field, value)` 返回普通编辑；对于完全相同的无操作输入返回空数组；对于无效或不受支持的候选值返回 `undefined`。自定义控件可以将这些辅助函数与现有的原子编辑 API 配合使用。数值窗格不支持带单位的字符串、公式输入以及更改销钉位置。

## 查找图表文本 {#find-diagram-text}

共享的搜索控件会在可见的形状和页面中查找字面文本。以编程方式使用的宿主可调用 `controller.setSearchQuery(query)`、`nextSearchResult()`、`previousSearchResult()` 和 `selectSearchResult(index)`；不可变的查询/结果可在 `controller.state.search` 中获得。查询不会改变选择，直到明确进行导航。新文档会清除搜索，而较新的重入宿主操作会取代旧的结果导航。

搜索对每个匹配的形状返回一条结果，并带有有上限的文本预览。它使用 ECMAScript 的小写比较，而不是正则表达式或感知语言的排序规则。限制为：查询 256 个字符、500 个匹配形状、每个形状 32,768 个已索引字符，总计 200 万个已索引字符。状态会明确指出部分结果。隐藏的子树和被抑制的组文本会被排除；背景页面作为独立页面仅索引一次。尚未实现逐处出现的高亮和形状数据搜索。

## 显示图层 {#display-layers}

使用共享的“图层”控件，或任何已挂载的/原生句柄的 `setLayerVisibility(pageId, layerId, visible)` 来覆盖显示可见性。`visible: null` 会恢复某个已保存的标志；`resetLayerVisibility(pageId?)` 会重置一个源页面或整个文档。源页面 ID 用于区分前景和背景图层。冻结的覆盖项可在 `controller.state.layerVisibilityOverrides` 中获得。

图层更改会同步渲染、选择和搜索，而不改变文档。参考线、NoShow 以及描述不完整的旧版隐藏形状无法通过图层覆盖显示出来。无障碍面板最多显示 200 个唯一图层，并会提示被截断；API 最多允许 25,000 个文档范围的覆盖项。SVG 和打印产物继续使用已保存的可见性。图层颜色、编辑以及已保存的打印策略均未实现。

## 功能区加载项选项卡 {#ribbon-add-in-tabs}

宿主可以在“帮助”之后添加选项卡，就像加载项在 Visio 中所做的那样。`ribbonAddIns` 是元素的属性，也是每个框架绑定的 prop，而 `office-ribbon-add-in` 是事件映射的一部分：

```tsx
const tabs = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [
			{
				label: 'Export',
				commands: [{ id: 'export', label: 'Export', icon: 'save', run: () => exportReport() }],
			},
		],
	},
];

<VisioViewer
	ribbonAddIns={tabs}
	events={{ 'office-ribbon-add-in': ({ tab, command }) => track(tab, command) }}
/>;
```

Vue 使用 `:ribbon-add-ins` 并发出 `office-ribbon-add-in`；Angular 使用 `[ribbonAddIns]` 并提供 `(ribbonAddIn)` 输出；原生绑定使用 `mountViewer(host, { ribbonAddIns })` 和 `viewer.update({ ribbonAddIns })`。对于元素本身，请设置 `element.ribbonAddIns` 并监听 DOM 事件。

命令会先运行其 `run` 回调，然后触发该事件。省略该 prop 会保留当前的选项卡；传入空数组会移除它们。再次传入相等的描述符（框架在每次渲染时都会这样做）会保留面板，只替换回调。该描述符（`RibbonAddInTab`，来自 `ooxml-core/ribbon`，同时由 `ooxml-ui` 和每个框架包导出）与 Word、Excel 和 PowerPoint 编辑器共用。选项卡不能使用内置选项卡的 id（`home`、`insert`、`design`、`data`、`process`、`review`、`view`、`help`）。

## 自定义 {#customization}

`customization` 接受一个 `VisioCustomization` 对象，即 `ooxml-core/embedding` 的共享约定，其他编辑器也使用它。每个绑定都将其作为 `customization` prop 或输入转发。对象标识变化时才会应用，因此请保持其稳定；`null` 恢复默认值。

```ts
const customization = {
	capabilities: { export: false, print: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [{ id: 'send', label: 'Send to chat', onSelect: ({ shapes }) => send(shapes) }],
	},
};
```

- `capabilities`：`false` 会移除通往该能力（`edit`, `open`, `save`, `export`, `print`, `share`, `settings`, `copy`, `comment`, `find`）的所有入口：功能区、“文件”选项卡页面、菜单和快捷键。`hiddenActions` 不变，两者取并集。
- `backstage.extraPages`：宿主在“文件”选项卡中添加的页面 `{ id, label, group?, order?, render }`，位于其分组的内置页面之后。`render` 是挂载函数、自定义元素标签或 `{ element, properties }`，并会收到一个 `VisioRegionContext`。
- `contextMenu.extraCommands`：宿主命令 `{ id, label, icon?, group?, when?, disabled?, onSelect }`，位于内置项之上（`top`，默认）或之下（`bottom`）。`onSelect` 在菜单关闭后运行，并收到一个 `VisioMenuContext`。标签按原样显示（请自行本地化）。

`edit: false` 是独立的只读锁，与宿主的只读开关和协作角色 `viewer` 并存：彼此互不解除。能力是界面策略，而不是文档保护。Visio 暂不提供区域替换（选项对话框、窗格）。各产品汇总：[嵌入指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md)。

## 静态 SVG 导出 {#static-svg-export}

`exportPageSvg(model, pageIndex, { maxBytes })` 以及每个已挂载的/原生句柄的 `exportSvg()` 都会返回当前页面的快照，而不改变文档、选择或视口。导出包括受支持的背景、嵌入的栅格资源，以及位于 SVG 描述/元数据中的可见兼容性说明。SVG 导出没有远程资源，也没有自动下载。演示仅在明确按下按钮后才会下载。

UTF-8 的上限为 16 MiB；调用方可以降低它。保守的预检可能会拒绝那些最终序列化结果实际更小的内容。字体不会被嵌入，文本/布局仍是近似的。物理页面尺寸会被保留；下游的栅格化程序必须在分配图像表面之前限制其输出尺寸和像素面积。PDF、位图导出和打印布局仍未实现。

可选的 Linux 次级像素测试在构建之后通过 `node scripts/test-svg-rasterization.mjs` 运行，使用系统 Python 3、librsvg/Cairo 以及核心的原生 canvas 依赖。这些检查与浏览器或 Microsoft Visio 验证是分开的。

## 准备绘图页面产物 {#prepare-drawing-page-artifacts}

`createPrintSnapshot(model, { pageIndices: [0, 2] })` 保留明确指定的页面顺序，并返回深度冻结的、相互独立的 SVG 记录，包含绘图尺寸和全部导出诊断信息。每个已挂载的/原生句柄的 `createPrintSnapshot()` 会捕获其当前页面。在准备期间如果文档被替换，它会拒绝，且不改变视图状态。不存在框架、对话框、下载或打印机命令。

快照使用已保存的显示外观。图层的“打印”、NonPrinting 以及已保存的打印机设置不会被应用；绘图尺寸不等于打印纸张或打印比例。结果中包含这些限制说明。保护措施限制了所选页面数、聚合字节数、重复的背景、栅格尺寸、文本/路径工作量以及重复的整模型验证。调用方可以降低限制。各个独立的 SVG 必须保持隔离，因为它们内部的资源 ID 可能重复。这是为将来的打印工作流所做的准备，而不是对原生 Visio 打印的支持。

## 限定范围的查找与替换 {#scoped-find-and-replace}

“开始 > 查找 > 替换”和 Ctrl+H 会打开共享的替换控件。“查找下一个”“上一个”“替换”和“全部替换”使用完整的、字面的、区分大小写的匹配项。范围为“所选内容”“当前页面”和“所有页面”。“所选内容”包含被选中的组子树，按选择顺序/源顺序排列，重叠的目标会被去重。页面范围使用已保存的源顺序，包括隐藏的源文本。导航会循环回绕，并在可能时选中可见的目标。区分大小写已勾选，且目前是固定的；全字匹配、大小写折叠和特殊搜索代码均未实现。这些范围/遍历规则是有上限的应用策略，并不声称与原生“查找”对话框等价。

每个原生绑定都暴露相同的 `controller`：宿主可以使用 `captureTextReplaceToken(scope)`、`getTextReplaceOccurrences(token, query)`、`selectTextReplaceOccurrence(token, occurrence, query)`、`planTextReplacement(token, { query, replacement, mode, current? })` 和 `applyTextReplacePlan(plan, token)`。导航和应用都会返回一个新的不透明令牌。保留该令牌可以在匹配项之间移动时保持原先选定的范围。手动更改选择/页面、替换源或进行其他操作都会使之前的令牌失效。`cancelTextReplace(token)` 只会取消该令牌正在进行的应用，而不影响较新的无关编辑继续运行。计划和令牌属于其所在的控制器，不能被伪造或跨实例复用。

替换使用一次原子的源编辑和一个撤销边界，包括跨页面的结果导航。原生的 `document-change`、`selection-change` 和 `shape-select` 事件仍可通过框架回调获得。“全部替换”会保留选择；替换当前项会前进到下一个匹配项。相同文本的替换会保持源字节不变且不增加历史记录，但仍会进行导航。不受支持的匹配到的富文本、含字段的文本、与母版关联的文本、受保护的文本或有依赖的文本，会使整批操作被拒绝。草稿会显示错误并保持可编辑。关闭控件或更改意图会取消其待处理的操作。

## 实验性本地纯文本编辑 {#experimental-local-plain-text-editing}

在 `load(bytesOrBlob)` 之后，选择一个本地形状并使用共享的文本控件，或者在任何已挂载的/原生句柄上调用 `replacePlainText(pageId, shapeId, text)`。`undo()` 和 `redo()` 使用有上限的文档历史；`cancelEdit()` 会取消待处理的工作。六个适配器都转发相同的方法和 `document-change` 事件，其 detail 为 `{ document, dirty, kind }`（kind 为 `edit`、`undo` 或 `redo`，另有 `remote` 和 `save`；见上文）。检查 `controller.state.edit` 可了解源是否可用、忙碌/脏状态、撤销/重做是否可用、历史是否被截断、错误和诊断信息。

由查看器加载或创建的、以源为依据的 VSDX 文档可以被编辑，或导出为 VSDX。二进制 VSD v11 导入通过 ooxml-core 使用已发布的 ole2 编解码器：受支持的显式页面/形状变换、移动/线条几何和纯文本会以带有诊断说明的回退样式进行预览。原始的旧版字节会被私下保留；旧版源的编辑、撤销/重做和 VSDX 导出均被禁用。更早的二进制版本、组/母版、复杂几何、模具、模板和启用宏的格式仍不受支持。通过 `document` 赋值的模型不提供可编辑的包字节。核心验证会拒绝与母版关联的形状、富文本、字段、已签名的包以及启用宏的内容。XML/包的编辑在隔离的 Worker 中运行；框架包装器不实现格式逻辑。纯文本编辑会拒绝受影响的或未知的、依赖文本的公式缓存，而不是让它们过期，并遵守有效的 `LockTextEdit` 保护。VSDX 的逻辑文本不包括其存储的末尾段落标记，同时保留有意的末尾 LF 字符和空白段落。旧版 VSD 的文本语义保持不变。富文本内容编辑仍不受支持。

`exportVsdx()` 返回 `{ bytes, dirty, diagnostics }`，用于明确下载的副本。它不会覆盖源文件、上传它，也不声称具备原生往返保真度。共享 UI 仅在用户明确操作之后才下载。历史记录有上限，可能会丢弃较早的撤销状态，这通过 `historyTruncated` 报告。纯文本替换仍限于上述已获准的子集。下文的格式设置和几何操作有各自独立的准入规则。原生重新打开的证据涵盖具体的已记录用例，详见[验证记录](/verification)；它并不证明通用保真度。

## 格式设置与排列 {#experimental-formatting-and-arrangement}

“开始”功能区和 `applyEdits()` 共用 `format-text`、`format-shape` 和 `reorder-shape` 命令。文本格式设置包括已保存的字体系列、磅值大小、粗体/斜体/下划线/删除线、颜色、项目符号、缩进以及水平/垂直对齐。形状格式设置包括纯色填充、线条颜色/粗细、线条图案 0-23、填充图案 0-24、背景色和透明度。格式设置要求是普通的本地叶子形状。针对整个形状的文本格式设置会更新每个有效的字符或段落行，同时保留运行标记、文本、无关的样式位以及受支持的缓存字段。字符格式和段落格式有各自独立的源准入；垂直对齐不会改写这两个部分中的任何一个。受保护的或有歧义的行以及受影响的公式依赖项会使整个事务被拒绝。字号增大和减小要求大小一致；字号选择器可以为大小不一的选区设置字号。当其祖先和缓存值/单位可以被证明时，本地的 Character/Paragraph `F="Inh"` 单元格是可编辑的。无缓存的委托需要真实的祖先；明确的格式选择会创建本地覆盖，即使值没有变化。顶层继承的保护和变换单元格仍受到保护。字体必须已存在于源绘图中。层叠命令支持对显示层带零中的一个普通顶层形状执行置于顶层/置于底层/上移一层/下移一层。

`format-shape` 接受 `linePattern`、`fillPattern`、`fillBackgroundColor`（`'#RRGGBB'`）、`lineTransparency` 和 `fillTransparency`（0 到 100 的百分比）。透明度按半个百分点的步长取整，恰好居中时向上取整。填充透明度会同时设置前景和背景。图案 0 表示不绘制；图案 1 为纯色。显式的 `linePattern: 1` 会独立于线条颜色启用轮廓。显式的填充图案会覆盖由 `fillColor` 暗含的纯色图案；`fillColor: 'none'` 与非零的填充图案同时使用会被拒绝。活动的渐变透明度需要显式的绘制替换，而普通的继承纯色主题绘制则受支持。

共享的“线条”菜单提供“无线条”和内置图案。两个绘制菜单中的“更多选项”以及形状右键菜单都会打开“设置形状格式”任务窗格，其中每项更改会立即作为一次选择事务应用。混合的源值在更改之前显示为空。`fillColorTheme`、`lineColorTheme` 以及（在 `format-text` 上的）`fontColorTheme` 指明了颜色来自哪个“主题颜色”色板，因此会保存为 Visio 的主题公式；`beginArrow`、`endArrow`、`beginArrowSize`、`endArrowSize`、`lineCap` 和 `rounding` 用于设置线条端点、端帽和圆角。六个原生绑定都使用现有的 `applyEdits()` 以及响应式的文档/选择事件；没有添加框架专用的绘制引擎。`visioShapeFormattingState(shapes)` 为自定义控件提供汇总的源绘制值。这些值与渐变/平铺的不透明度乘数是分开的。

六个“对齐”命令以所选的第一个形状为基准对齐。“位置”命令为三个或更多形状提供相等的水平和垂直边缘间距。核心的 `visioArrangeCommands(page, selectedIds, action)` 辅助函数为带有自定义控件的宿主生成原子编辑。它测量旋转后的宽/高边框，按中心点排序分布，并在并列时保留选择顺序。组、母版、图层和已粘合的连接线仍在其准入范围之外。这些操作与其他编辑共享历史记录、选择保留以及 VSDX 下载/重新打开。

## 复制 {#experimental-duplication}

控制器以及每个已挂载的/原生句柄上的 `duplicateSelection(): Promise<void>` 会复制当前页面上获准的选中形状。“开始”的“粘贴”菜单、形状右键菜单、“告诉我”和 Ctrl+D 调用同一个方法。文本输入框保留其原生键盘行为。复制不使用系统剪贴板。

核心的 `visioDuplicateCommand(page, selectedIds)` 辅助函数生成一个带有全新页面本地 ID 的 `duplicate-shapes` 源命令。副本保留源 XML、混合文本和层叠顺序，获得不含复制来的唯一 ID 的全新名称，并向右下方移动 0.33 绘图英寸。源验证目前接受本地的、未粘合的 2D 叶子形状，排除母版、组、外来内容、图层、有歧义的标识和不受支持的依赖项。在接受任何源更改之前，会检查销钉保护和受影响的公式缓存。

生成的副本在一个原子历史步骤中成为选区。撤销会恢复原来的选区，重做会恢复副本，除非更新的选择或页面操作取代了它。失败、被取消或被取代的操作不会部分地改变源字节或选择。普通的 `selection-change` 和 `document-change` 事件会向原生框架状态公开已接受的结果。

## 实验性形状剪贴板 {#experimental-shape-clipboard}

每个已挂载的/原生句柄都提供 `copySelection()`、`cutSelection()` 和 `pasteSelection()`，均返回 `Promise<void>`。请从用户手势中调用它们，以便浏览器授予剪贴板访问权限。“开始”、右键菜单、“告诉我”以及 Ctrl/Meta+C、X 和 V 使用同一个共享适配器。文本输入框保留其原生剪贴板行为。剪贴板失败会显示在查看器中，而不改变源字节、历史记录或选择。剪切会先写入剪贴板再删除；如果保护阻止删除，被复制的形状仍保留在剪贴板中，查看器会报告它们无法被剪切。

负载是带有 `OOXML-VISIO-SHAPES/1` 标记的、有上限的 `text/plain`。它包含捕获的形状 XML 和所需的源上下文，因此编辑或剪切原始形状不会使已复制的内容失效。粘贴始终读取当前的系统剪贴板；它不会悄悄地复用内部的上一次复制。原生 Microsoft Visio 剪贴板格式、任意文本、图像和选择性粘贴均不受支持。粘贴的形状会获得全新的 ID/名称，并向右下方偏移 0.33 绘图英寸。重复粘贴使用相同的偏移；原生的粘贴级联和跨页面放置尚未得到验证。

首个受支持的范围是普通的本地、未粘合的 2D 叶子形状，要求文档的字体/样式/主题定义以及页面设置、序号、背景上下文和绘图比例相匹配。引用必须保留在被复制的选区内。组、母版、图层、外来内容、含关系的片段、不受支持的公式以及资源导入仍会被拒绝。核心会在原子的源编辑之前验证整个负载。粘贴会选中其副本，并与“复制”共用有上限的历史记录和选择意图规则。

`controller.state.clipboard` 通过各框架现有的响应式订阅公开不可变的 `ready`、`preparing` 和 `error` 状态。源捕获在可取消的 Worker 中运行，并会因选择、页面、可见性或文档的更改而失效。提供自己的剪贴板传输的宿主可以在控制器上使用 `captureClipboardToken()`、`prepareClipboardSelection(token)`、`getPreparedClipboard(token)`、`cutPreparedSelection(token)` 和 `pasteClipboardText(text, token)`。令牌会拒绝过期的操作。不依赖 DOM 的核心导出 `captureVisioClipboard`、`serializeVisioClipboard`、`deserializeVisioClipboard`、`paste-shapes` 编辑以及 `visioPasteCommand(page, clipboard, offset?)` 命令构建器。

## 几何编辑 {#experimental-geometry-editing}

共享的几何控件提供创建矩形、移动、调整大小和安全删除。每个原生/已挂载的框架句柄还转发 `applyEdits(edits: readonly VisioEdit[])`，用于通过相同的 Worker、历史记录和取消路径执行原子批处理。

“开始 > 文本工具”和“插入 > 文本框”使用同一个共享的固定大小文本框草稿。拖动页面空白处，输入纯文本，然后使用“添加文本框”或 Ctrl/Meta+Enter；按 Escape 取消。空草稿不会创建源形状；有意输入的空格和换行会被保留。使用文本工具点击现有形状会选中它，并显示现有的文本编辑控件。矩形、椭圆和线条的绘制使用相同的受保护创建生命周期。其内部的控制器事务在绘制之前捕获源、页面、有序选择、缩放和可见性意图，在接受源之前检查该意图，并在同一个撤销步骤中选中创建的形状。更新的意图或销毁会取代待处理的创建。内部的创建令牌不是额外的原生绑定 API。

所有六个公共句柄都通过 `applyEdits()` 接受这个新的普通命令：

```ts
await handle.applyEdits([
	{
		type: 'create-text-box',
		pageId: '0',
		shapeId: '10',
		x: 4,
		y: 5,
		width: 3,
		height: 2,
		text: 'Text with an intentional trailing blank paragraph\n',
	},
]);
```

文本是必需的。已保存的默认文本样式提供初始的字符格式。显式的 `FillPattern=0` 和 `LinePattern=0` 使新文本框不绘制任何内容；仅更改线条颜色不会启用轮廓。固定大小的草稿创建并不代表原生 GUI 文本工具的默认值、自动文本大小调整或富文本内容编辑。通用的 `applyEdits()` 保持其现有的选择行为；共享的绘图工具负责创建后形状的选择。原生证据涵盖三种绘图比例下 24 个已保存的固定文本框 API 用例，使用默认的和自定义的“仅文本”/Arial 18 pt 样式，核心输出已在 Visio 中重新打开。这些已保存样式的用例并未衡量交互式 GUI 的默认值；参见[验证记录](/verification)。

“插入 > 新建页面”会经由现有的、以源为依据的页面插入命令，然后选中新页面。它与页面栏的插入行为、撤销/重做和源准入共用同一套机制；这并不增加基于模板的页面创建。

```ts
await handle.applyEdits([
	{
		type: 'create-rectangle',
		pageId: '0',
		shapeId: '9',
		x: 4,
		y: 5,
		width: 3,
		height: 2,
		text: 'New shape',
	},
	{ type: 'move-shape', pageId: '0', shapeId: '1', x: 6, y: 7 },
]);
```

坐标是以绘图英寸为单位的旋转销钉位置，原点位于左下角，向上为正。除非命令包含 `anchor: { x, y }`，否则调整大小会固定销钉。每个锚点坐标为 `0`、`0.5` 或 `1`，用于标识形状规范化的、向上的本地宽/高边界中的固定点。例如，`{ x: 0, y: 0 }` 会固定其本地左下角。锚定的调整大小对本地销钉、受影响的缓存以及固定的旋转/翻转值有独立的源证明。ID 是显式指定的。现有的、获准的本地顶层 2D 形状可以被编辑，包括来自其他生成器的形状。核心会对受影响的、受支持的数值 ShapeSheet 依赖项求值，并保留未被触及的 ZIP 负载。查看器绝不自行修改 XML。

这是一个狭窄的、已获准的子集。母版/组/外来形状、连接线/粘合、不安全的保护/重定向、不受支持的受影响公式、有歧义的包依赖项以及被引用的删除都会失败并显示可见的错误。相对线条几何会缩放；绝对线条几何需要受支持的尺寸依赖。命令约定、数值限制、排除原因及后续扩展见核心的 [`src/core/visio/README.md`](https://github.com/ChristopherVR/ooxml/blob/main/src/core/visio/README.md)。[功能清单](/parity)记录了当前获准的集合，现在它还包括带有显式本地销钉的、有上限的基本母版实例移动。

生成的/导入的简单形状测试以及本地 Chrome 工作流均已通过。目前已接受的 19 个公开语料库图表获准的几何编辑数为零，因为尚无法证明非页面依赖范围是相互独立的。其中十三个还缺少符合条件的本地形状。该语料库的实际编辑覆盖仍受阻；后续工作是建立限定范围的包/母版/主题依赖图和母版实例编辑，然后是粘合端点的路由。此后，针对[验证记录](/verification)中特定格式设置和排列的捕获，已验证了原生重新打开；通用保真度仍未得到验证。
