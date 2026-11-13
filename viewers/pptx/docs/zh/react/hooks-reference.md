---
title: 完整 Hooks 参考
description: 按职责列出组成 PowerPointViewer 的全部内部 hook，以及从 pptx-react-viewer/internals 导入的方式。
---

# 完整 Hooks 参考 {#complete-hooks-reference}

本页是 [Hooks](/zh/react/hooks) 引用的完整列表，按职责组织 `PowerPointViewer` 的内部 hooks。遵循语义化版本兼容保证的公共子集见[公共 hooks](/zh/react/hooks#public-hooks)。

::: warning 内部构建块
下列 hooks 均可从 `pptx-react-viewer/internals` 导入，但不受语义化版本兼容保证约束，签名、行为和是否存在都可能在不升级主版本的情况下改变。只有 `pptx-react-viewer/viewer` 的公共导出和组件的属性、句柄不能满足需求时才应使用，并锁定精确版本。

```tsx
import { useViewerState, useEditorHistory } from 'pptx-react-viewer/internals';
```

:::

## 核心状态与生命周期 {#core-state-lifecycle}

| Hook                     | 职责                                        |
| ------------------------ | ------------------------------------------- |
| `useViewerState`         | 组合核心和 UI 状态。                        |
| `useViewerCoreState`     | 幻灯片、选择、画布尺寸和模式。              |
| `useViewerUIState`       | 面板可见性、对话框和工具栏标记。            |
| `useDerivedSlideState`   | 计算可见索引、节和母版伪幻灯片。            |
| `useDerivedElementState` | 计算元素边界、控点和堆叠顺序等派生状态。    |
| `useContentLifecycle`    | 连接内容同步、未保存修改跟踪和恢复检测。    |
| `useLoadContent`         | 挂载时通过 `PptxHandler` 解析 PPTX 缓冲区。 |
| `useRecoveryDetection`   | 检测之前未保存的会话，提供恢复。            |
| `useAutosave`            | 定时自动保存调度和状态。                    |
| `useSerialize`           | 将当前文档序列化为字节。                    |
| `useViewerIntegration`   | 顶层整合 I/O、导出、打印、指针和生命周期。  |

## 编辑与历史记录 {#editing-history}

| Hook                         | 职责                                       |
| ---------------------------- | ------------------------------------------ |
| `useEditorHistory`           | 撤销重做快照栈，指针交互期间延迟捕获。     |
| `useEditorOperations`        | 组合全部编辑操作。                         |
| `useElementOperations`       | 元素创建、更新和删除的基础操作。           |
| `useElementManipulation`     | 移动、缩放、旋转和删除元素。               |
| `useSectionOperations`       | 创建、重命名和删除节。                     |
| `useTableOperations`         | 插入和删除行列，合并和拆分单元格。         |
| `useSlideManagement`         | 添加、删除、复制、重排和隐藏幻灯片。       |
| `useInsertElements`          | 插入形状、图片、文本框、表格和图表。       |
| `useGroupAlignLayerHandlers` | 组合、取消组合、对齐、分布和前后图层操作。 |
| `useMergeShapesHandler`      | 形状的联合、相减、相交和排除等布尔运算。   |
| `usePropertyHandlers`        | 选中元素的属性面板变更处理。               |
| `useThemeHandlers`           | 应用和切换文档的 PowerPoint 主题。         |
| `useThemeSwitching`          | 更高层的主题切换协调，也属于公共导出。     |
| `useLayoutSwitching`         | 切换幻灯片版式并重新映射占位符。           |
| `useClipboardHandlers`       | 元素复制、剪切和粘贴。                     |
| `useFindReplace`             | 跨幻灯片文本查找替换。                     |
| `useComments`                | 批注讨论串状态和增删改查。                 |
| `useAnnotationHandlers`      | 自由绘制和形状批注处理。                   |

## 画布交互 {#canvas-interaction}

| Hook                        | 职责                                |
| --------------------------- | ----------------------------------- |
| `usePointerHandlers`        | 画布鼠标和触控事件。                |
| `useCanvasInteractions`     | 选框、拖动选择和画布手势。          |
| `useZoomViewport`           | 缩放、适应宽度和视口 DOM ref。      |
| `useKeyboardShortcuts`      | 快捷键定义。                        |
| `useKeyboardShortcutWiring` | 将快捷键定义连接到 DOM 事件监听器。 |
| `useResizablePanels`        | 调整属性面板和侧边栏宽度。          |
| `useSheetDismissDrag`       | 移动端底部抽屉滑动关闭。            |
| `useModalDismissDrag`       | 移动端对话框滑动关闭。              |

## 导出、打印与 I/O {#export-print-i-o}

| Hook                                               | 职责                                   |
| -------------------------------------------------- | -------------------------------------- |
| `useExportHandlers`                                | PNG、SVG、PDF、GIF、视频和 PPTX 导出。 |
| `useExportSaveAs`                                  | 导出时的另存为文件选择流程。           |
| `usePrintHandlers`                                 | 打印对话框和布局。                     |
| `useIOHandlers`                                    | 打开和导入文件。                       |
| `useFontInjection`                                 | 将文稿嵌入字体注入页面。               |
| `useVirtualizedSlides`（含 `computeVirtualRange`） | 为大型文稿提供幻灯片面板和列表虚拟化。 |

## 对话框 {#dialogs}

| Hook                   | 职责                         |
| ---------------------- | ---------------------------- |
| `useViewerDialogs`     | 各模态对话框的开关状态。     |
| `useDialogCustomShows` | 特定对话框的自定义显示条件。 |

## 放映模式 {#presentation-mode}

| Hook                         | 职责                             |
| ---------------------------- | -------------------------------- |
| `usePresentationMode`        | 顶层组合放映导航、动画和切换。   |
| `usePresentationSetup`       | 进入放映时设置全屏和初始幻灯片。 |
| `usePresentationAnnotations` | 放映时的画笔和荧光笔批注。       |
| `useAnimationPlayback`       | 播放元素动画时间线。             |
| `useRehearsalTimings`        | 排练时记录各页时长。             |
| `usePresentationKeyboard`    | 放映期间的键盘导航。             |
| `useSlideNavigation`         | 上一页、下一页和跳转。           |
| `useZoomNavigation`          | 放映期间缩放和平移。             |
| `usePresenterWindow`         | 演讲者窗口与观众窗口消息通信。   |
| `useAudienceMode`            | 观众窗口的渲染模式。             |

## 幻灯片过渡辅助函数 {#slide-transition-helpers}

所有绑定共用的框架无关逻辑位于 `ooxml-ui/pptx`；私有的 `pptx-viewer-shared` 包只是它的兼容外观。如果宿主搭建了自己的放映舞台（自定义 `SlideStage`，而非完整的 `PowerPointViewer`），仍然需要用到过渡解析器和关键帧，因此这里把整套接口重新导出：

| 导出项                                                                                                                                                          | 职责                                                                                          |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| `resolveSlideTransition`                                                                                                                                        | 将 `PptxSlideTransition` 解析为退出层/进入层的 CSS `animation` 简写属性。                     |
| `resolveTransitionDurationMs`                                                                                                                                   | 计算过渡的有效时长（毫秒），会考虑手动设置的时长、旧版 `spd` 取值和 PowerPoint 自身的默认值。 |
| `getSlideTransitionAnimations`                                                                                                                                  | `resolveSlideTransition` 内部调用的底层解析函数（经典二维过渡系列）。                         |
| `getCinematicTransitionAnimations` / `getP14TransitionAnimations`                                                                                               | 分别对应 Office 2013+（p15）影院级过渡系列和 Office 2010（p14）特效/三维过渡系列。            |
| `SLIDE_TRANSITION_KEYFRAMES`（别名 `SLIDE_TRANSITION_KEYFRAMES_CSS`）                                                                                           | 上述所有解析出的动画名称所引用的完整 `@keyframes` 代码块。只需通过 `<style>` 标签注入一次。   |
| `CINEMATIC_TRANSITION_KEYFRAMES` / `P14_TRANSITION_KEYFRAMES_ALL`                                                                                               | 已经并入 `SLIDE_TRANSITION_KEYFRAMES` 的关键帧子代码块；同时也单独导出。                      |
| `resolveDirection` / `resolveDirection8` / `resolveOrientation`                                                                                                 | 将 OOXML 的 `dir` / `orient` 取值归一化为已解析的方向/朝向。                                  |
| `resolveWheelSpokeCount`                                                                                                                                        | 将手动设置的 Wheel（车轮）辐条数吸附到 PowerPoint 提供的最接近取值。                          |
| `RANDOM_ELIGIBLE_TYPES`、`INSTANT`、`DEFAULT_TRANSITION_DURATION_MS`、`DEFAULT_MORPH_DURATION_MS`、`TRANSITION_SPEED_DURATION_MS`、`EASE`、`WHEEL_SPOKE_COUNTS` | 上述解析函数使用的辅助常量。                                                                  |
| `PresentationTransitionOverlay`、`MorphTransitionOverlay`、`SlideLayer`、`FragmentedTransitionLayer`                                                            | 放映模式过渡叠加层组件本身，供需要整个渲染叠加层（而不只是 CSS 解析结果）的宿主使用。         |

## 协作 {#collaboration}

公共子集见[协作](/zh/react/collaboration)，完整内部集合如下：

| Hook                      | 职责                                   |
| ------------------------- | -------------------------------------- |
| `useYjsProvider`          | 管理 Yjs WebSocket Provider 生命周期。 |
| `useYjsDocumentSync`      | 将 Yjs 文档同步到预览器状态。          |
| `usePresenceTracking`     | 跟踪远程光标、选择和连接状态。         |
| `useCollaborativeState`   | 基于 CRDT 的共享文档状态。             |
| `useCollaborativeHistory` | 协作撤销重做。                         |
| `useBroadcastFollower`    | 跟随其他用户广播的视口和选择。         |
| `useFollowMode`           | 协调跟随演讲者模式。                   |

## 移动端与响应式 {#mobile-responsive}

| Hook                 | 职责                                      |
| -------------------- | ----------------------------------------- |
| `useIsMobile`        | 判断设备、视口、手机、平板和方向。        |
| `useTouchGestures`   | 识别双指缩放、平移和点击手势。            |
| `useKeyboardInsets`  | 根据移动端屏幕键盘调整布局。              |
| `useSwipeNavigation` | 移动端滑动翻页。                          |
| `useReducedMotion`   | 动画和切换遵循 `prefers-reduced-motion`。 |

列表来自 `packages/react/src/viewer/hooks/**/*.ts`，全部从 [`pptx-react-viewer/internals`](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/react/src/internals.ts) 重新导出。在该目录新增或重命名 hook 时，请同步更新本页。

## useViewerBuildingBlocks：宿主工作流 {#building-block-host-workflows}

`useViewerBuildingBlocks` 从包根入口（`pptx-react-viewer`）导出，而不仅从 `pptx-react-viewer/internals` 导出。
它的 `actions` 对象用于完成构建块不挂载的选择器、对话框和版式工作流。使用说明和示例见
[Hooks](/zh/react/hooks)，本节列出精确的形状。

### 选项 {#building-block-options}

| 选项         | 类型                                      | 说明                                                                                                      |
| ------------ | ----------------------------------------- | --------------------------------------------------------------------------------------------------------- |
| `onRequest`  | `(request: BuildingBlockRequest) => void` | 工具栏控件或画布触发了构建块无法挂载的请求时调用。该请求同时保存在 `actions.pending` 中，直到完成或取消。 |
| `fileLimits` | `HostFileLimits`                          | 宿主自有选择器所选文件的上限。超出上限返回 `'file-too-large'`。默认图片 25 MiB，音频和视频 100 MiB。      |

```ts
interface HostFileLimits {
	/** Largest accepted image, in bytes. Defaults to 25 MiB (26214400). */
	imageBytes?: number;
	/** Largest accepted audio or video file, in bytes. Defaults to 100 MiB (104857600). */
	mediaBytes?: number;
}
```

### 请求 {#building-block-requests}

```ts
type BuildingBlockRequest =
	| { kind: 'image' }
	| { kind: 'media' }
	| { kind: 'hyperlink' }
	| { kind: 'smartArt' }
	| { kind: 'chart' }
	| { kind: 'equation'; existingOmml: Record<string, unknown> | null };
```

| `kind`      | 触发方                                                                             | 完成方式                                                                            | 可用条件                           |
| ----------- | ---------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------- | ---------------------------------- |
| `image`     | 图片控件                                                                           | `actions.insertImage(file)`                                                         | `insertImage`                      |
| `media`     | 视频控件                                                                           | `actions.insertMedia(file)`                                                         | `insertMedia`                      |
| `hyperlink` | 链接控件及其画布手势                                                               | `actions.hyperlink.apply(data)`                                                     | `hyperlink`                        |
| `smartArt`  | SmartArt 控件                                                                      | `actions.smartArt.insert(layout, items)`                                            | `insertSmartArt`                   |
| `chart`     | 图表控件                                                                           | `actions.chart.insert(kind, options?)`                                              | `insertChart`                      |
| `equation`  | 公式控件，或双击画布上的公式。`existingOmml` 为正在编辑的公式，新建公式时为 `null` | `actions.equation.insert(omml)`，编辑已有公式时使用 `actions.equation.update(omml)` | `insertEquation` 或 `editEquation` |

只有对应操作可用时才会触发请求。对应操作成功后会清除 `actions.pending`；操作失败则保留。
`actions.cancel()` 清除该请求，且不改变文档。

### 操作 {#building-block-operations}

操作返回 `BuildingBlockResult`（见[结果与失败代码](#building-block-results)），签名另有说明的除外。

| 成员                     | 签名                                                                           | 可用性标识                   |
| ------------------------ | ------------------------------------------------------------------------------ | ---------------------------- |
| `pending`                | `BuildingBlockRequest \| null`                                                 | 无                           |
| `cancel`                 | `() => void`                                                                   | 无                           |
| `insertImage`            | `(file: File) => Promise<BuildingBlockResult>`                                 | `insertImage`                |
| `insertMedia`            | `(file: File) => Promise<BuildingBlockResult>`                                 | `insertMedia`                |
| `hyperlink.current`      | `HyperlinkEditData \| null`                                                    | 无（读取选区）               |
| `hyperlink.slideCount`   | `number`                                                                       | 无                           |
| `hyperlink.apply`        | `(data: HyperlinkEditData) => BuildingBlockResult`                             | `hyperlink`                  |
| `hyperlink.remove`       | `() => BuildingBlockResult`                                                    | `removeHyperlink`            |
| `smartArt.insert`        | `(layout: SmartArtLayout, items: string[]) => BuildingBlockResult`             | `insertSmartArt`             |
| `chart.insert`           | `(kind: InsertChartKind, options?: InsertChartOptions) => BuildingBlockResult` | `insertChart`                |
| `equation.current`       | `Record<string, unknown> \| null`                                              | 无（读取选区）               |
| `equation.insert`        | `(omml: Record<string, unknown>) => BuildingBlockResult`                       | `insertEquation`             |
| `equation.update`        | `(omml: Record<string, unknown>) => BuildingBlockResult`                       | `editEquation`               |
| `findReplace`            | `BuildingBlockFindReplace`（见下文）                                           | `replace`（仅替换）          |
| `layout`                 | `BuildingBlockLayoutActions`（见下文）                                         | `applyLayout`、`resetLayout` |
| `availability`           | `Record<BuildingBlockActionId, BuildingBlockActionState>`                      | 无                           |
| `unmountedDialogActions` | `readonly string[]`                                                            | 无                           |

`hyperlink.slideCount` 是幻灯片链接选择器可链接的幻灯片数量。`hyperlink.current` 是选中元素的链接，
以编辑器值表示。`equation.current` 是选中元素的公式（OMML），没有则为 `null`。`insertImage` 会把图片
居中放置在幻灯片上并适配其尺寸。

`unmountedDialogActions` 列出构建块不挂载对话框的 `toolbarProps` 处理函数：`onOpenPasswordProtection`、
`onOpenDocumentProperties`、`onOpenFontEmbedding`、`onOpenDigitalSignatures`、`onOpenSetUpSlideShow`、
`onOpenCustomShows`、`onOpenBroadcastDialog` 和 `onToggleShortcuts`。可用 `hiddenActions` 隐藏它们，
或自行渲染对话框。

#### 查找和替换 {#building-block-find-replace}

`actions.findReplace` 暴露查看器的查找状态和操作。只读查看器中仍可查找；替换由 `replace` 可用性标识控制。

| 成员                            | 类型                                                                   | 说明                                                                                      |
| ------------------------------- | ---------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| `open`、`setOpen`               | `boolean`、`(open: boolean) => void`                                   | 查找栏是否打开。Ctrl+F 和 Ctrl+H 切换它。查找栏由宿主渲染。                               |
| `query`、`setQuery`             | `string`、`(query: string) => void`                                    |                                                                                           |
| `replacement`、`setReplacement` | `string`、`(replacement: string) => void`                              |                                                                                           |
| `options`、`setOptions`         | `FindReplaceOptions`、`(options: Partial<FindReplaceOptions>) => void` | `matchCase`、`includeNotes`、`includeChartTitles` 和 `includeChartData`，均为 `boolean`。 |
| `matches`                       | `readonly FindResult[]`                                                |                                                                                           |
| `currentIndex`                  | `number`                                                               | 从零开始。没有当前匹配时为 `-1`。                                                         |
| `currentMatch`                  | `FindResult \| null`                                                   |                                                                                           |
| `find`                          | `() => void`                                                           | 按当前查询和选项重新查找。                                                                |
| `next`、`previous`              | `() => void`                                                           | 显示匹配项所在的幻灯片并选中其元素。                                                      |
| `replaceCurrent`、`replaceAll`  | `() => BuildingBlockResult`                                            | 适用编辑权限、校验和历史记录。没有可替换内容时返回 `'no-match'`。                         |

#### 版式 {#building-block-layouts}

`actions.layout` 以活动幻灯片母版的版式作为允许集合。工具栏中的版式下拉和重置执行相同操作。

| 成员          | 类型                                             | 说明                                                           |
| ------------- | ------------------------------------------------ | -------------------------------------------------------------- |
| `current`     | `PptxLayoutOption \| undefined`                  | 活动幻灯片记录的版式（当文档列出该版式时）。                   |
| `currentPath` | `string \| undefined`                            |                                                                |
| `available`   | `PptxLayoutOption[]`                             | 活动幻灯片可切换到的版式。                                     |
| `apply`       | `(path: string) => Promise<BuildingBlockResult>` | 路径不在 `available` 中，或应用失败时返回 `'unknown-layout'`。 |
| `reset`       | `() => Promise<BuildingBlockResult>`             | 重新应用幻灯片自身的版式。未记录版式时返回 `'no-layout'`。     |

### 图表 {#building-block-charts}

```ts
type InsertChartKind =
	| 'column'
	| 'bar'
	| 'line'
	| 'pie'
	| 'doughnut'
	| 'area'
	| 'scatter'
	| 'histogram'
	| 'pareto'
	| 'funnel'
	| 'treemap'
	| 'sunburst'
	| 'boxWhisker'
	| 'regionMap';

type InsertChartGrouping = 'clustered' | 'stacked' | 'percentStacked';

interface InsertChartOptions {
	/** Default sample title when undefined; no title when empty. */
	title?: string;
	/** Only for the `column`, `bar`, `line` and `area` kinds. */
	grouping?: InsertChartGrouping;
	/** Defaults to true. */
	hasLegend?: boolean;
	/** Host data in place of the sample data. */
	data?: InsertChartData;
}

interface InsertChartData {
	categories: readonly string[];
	series: ReadonlyArray<{ name: string; values: readonly number[] }>;
}
```

`data` 至少需要一个分类和一个系列，且每个系列对每个分类都要有一个有限数值。否则 `chart.insert` 返回
`'invalid-data'`（在可用性检查之后）。

### 可用性 {#building-block-availability}

```ts
type BuildingBlockActionId =
	| 'insertImage'
	| 'insertMedia'
	| 'hyperlink'
	| 'removeHyperlink'
	| 'insertSmartArt'
	| 'insertChart'
	| 'insertEquation'
	| 'editEquation'
	| 'replace'
	| 'applyLayout'
	| 'resetLayout';

interface BuildingBlockActionState {
	/** The operation would run now. When false, `reason` says why. */
	available: boolean;
	reason?: BuildingBlockFailureCode;
	/** The toolbar control only raises a request; the host must render the picker or dialog. */
	needsHostUi: boolean;
}
```

每一项都先检查共同条件：`'not-ready'`（正在加载或加载失败）、`'read-only'`（非编辑模式或没有编辑权限）和
`'protected'`（建议只读或处于受保护视图）。之后检查该操作自身的条件。第一个不满足的条件即为 `reason`。

| 标识              | `needsHostUi` | 自身条件（按顺序）                                       |
| ----------------- | ------------- | -------------------------------------------------------- |
| `insertImage`     | true          | `'no-slide'`                                             |
| `insertMedia`     | true          | `'no-slide'`                                             |
| `hyperlink`       | true          | `'no-selection'`                                         |
| `removeHyperlink` | false         | `'no-selection'`，然后 `'no-hyperlink'`                  |
| `insertSmartArt`  | true          | `'no-slide'`                                             |
| `insertChart`     | true          | `'no-slide'`                                             |
| `insertEquation`  | true          | `'no-slide'`                                             |
| `editEquation`    | true          | `'no-selection'`，然后 `'no-equation'`                   |
| `replace`         | false         | 无                                                       |
| `applyLayout`     | false         | `'no-slide'`，然后 `'no-layout'`（无处理器或无版式）     |
| `resetLayout`     | false         | `'no-slide'`，然后 `'no-layout'`（无处理器或未记录版式） |

### 结果与失败代码 {#building-block-results}

```ts
type BuildingBlockResult = { ok: true } | { ok: false; code: BuildingBlockFailureCode };
```

失败代码不是英文文本，请用你自己的语言呈现。操作先报告可用性原因，再检查参数。失败的操作不会改变文档。

| `code`             | 含义                                                                         | 返回于                                                                                                             |
| ------------------ | ---------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| `not-ready`        | 文档仍在加载或加载失败。                                                     | 所有操作                                                                                                           |
| `read-only`        | 不可编辑：仅查看、非编辑模式或没有编辑权限。                                 | 所有操作                                                                                                           |
| `protected`        | 建议只读或处于受保护视图，因此禁止插入。                                     | 所有操作                                                                                                           |
| `no-slide`         | 没有活动幻灯片。                                                             | `insertImage`、`insertMedia`、`smartArt.insert`、`chart.insert`、`equation.insert`、`layout.apply`、`layout.reset` |
| `no-selection`     | 未选中任何内容。                                                             | `hyperlink.apply`、`hyperlink.remove`、`equation.update`                                                           |
| `no-file`          | 未提供文件。                                                                 | `insertImage`、`insertMedia`                                                                                       |
| `empty-file`       | 文件没有内容。                                                               | `insertImage`、`insertMedia`                                                                                       |
| `unsupported-type` | 图片操作收到的不是图片，或媒体操作收到的不是音频或视频类型。                 | `insertImage`、`insertMedia`                                                                                       |
| `file-too-large`   | 文件超过 `fileLimits` 的上限。                                               | `insertImage`、`insertMedia`                                                                                       |
| `decode-failed`    | 图片或媒体文件无法读取或解码。                                               | `insertImage`、`insertMedia`                                                                                       |
| `stale-target`     | 解码或探测文件期间，文档、活动幻灯片或其位置发生了变化。未插入任何内容。     | `insertImage`、`insertMedia`                                                                                       |
| `unsafe-url`       | 链接使用了脚本类协议（`javascript:`、`data:`、`vbscript:` 或 `mhtml:`）。    | `hyperlink.apply`                                                                                                  |
| `invalid-target`   | 链接既没有 URL 也没有操作。                                                  | `hyperlink.apply`                                                                                                  |
| `no-hyperlink`     | 选中元素没有链接。                                                           | `hyperlink.remove`                                                                                                 |
| `no-equation`      | 选中元素没有公式。                                                           | `equation.update`                                                                                                  |
| `no-match`         | 没有当前匹配项（`replaceCurrent`），或完全没有匹配项（`replaceAll`）。       | `findReplace.replaceCurrent`、`findReplace.replaceAll`                                                             |
| `unknown-layout`   | 路径不在 `layout.available` 中，或应用版式失败。                             | `layout.apply`                                                                                                     |
| `invalid-data`     | 传给 `chart.insert` 的 `data` 不符合[图表](#building-block-charts)中的规则。 | `chart.insert`                                                                                                     |
| `no-layout`        | 幻灯片未记录版式（重置），或无处理器或无版式（应用）。                       | `layout.apply`、`layout.reset`                                                                                     |
