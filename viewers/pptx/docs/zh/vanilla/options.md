---
title: 选项与回调
description: PptxViewerOptions 和 PptxViewerCallbacks 完整参考，涵盖源内容、主题、语言、界面开关、渲染器注册表、自动保存、协作，以及各个 onLoad/onError/onSlideChange 等回调。
---

# 选项与回调 {#options-callbacks}

`collaboration.externalSession` 选项可以复用应用自己持有的 Yjs 文档和 provider。参见[宿主自管协作](/zh/guide/host-owned-collaboration)。

自定义宿主尺寸时，请参阅[视口适配](/zh/guide/viewport-fit)，了解 `fitPadding`、`maxFitScale`、各框架示例和默认值。

`createPptxViewer(container, options)` 接受下列 `PptxViewerOptions` 接口。每个选项都可选，包括 `source`：省略时以空状态启动，稍后调用 [`loadFile` / `loadUrl`](/zh/vanilla/api#loading)。此参考直接依据 `packages/vanilla/src/viewer/types.ts` 整理。

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';
import type { PptxViewerOptions, PptxViewerCallbacks } from 'pptx-vanilla-viewer';
```

::: tip 提示
工厂函数还返回命令式句柄，参见[查看器实例 API](/zh/vanilla/api)，它不属于 `PptxViewerOptions`。
:::

## 内容 {#content}

| 选项       | 类型                 | 默认值 | 说明                                                                                                 |
| ---------- | -------------------- | ------ | ---------------------------------------------------------------------------------------------------- |
| `source`   | `PptxViewerSource`   | -      | 要打开的演示文稿：原始字节（`ArrayBuffer` / `Uint8Array`）、`Blob` / `File`，或要获取的 URL 字符串。 |
| `fileName` | `string`             | -      | PowerPoint 风格标题栏中显示的名称。                                                                  |
| `fonts`    | `ViewerFontSource[]` | -      | 宿主应用提供的已授权字体来源（`{ family, src, format?, weight?, style? }`）。                        |

```ts
type PptxViewerSource = ArrayBuffer | Uint8Array | Blob | string;
```

## 界面与初始状态 {#chrome-initial-state}

| 选项                      | 类型                | 默认值     | 说明                                                                                                                                                           |
| ------------------------- | ------------------- | ---------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `initialSlide`            | `number`            | `0`        | 加载后显示的幻灯片，索引从 0 开始，自动限制在有效范围内。                                                                                                      |
| `showToolbar`             | `boolean`           | `true`     | 显示导航、缩放和全屏工具栏。                                                                                                                                   |
| `showThumbnails`          | `boolean`           | `true`     | 显示缩略图侧边栏。                                                                                                                                             |
| `showCompatibilityToasts` | `boolean`           | `true`     | 加载后显示兼容性警告提示堆栈。`false` 只隐藏堆栈，警告仍会被收集。不受 `showToolbar` 影响。                                                                    |
| `showFormatToolbar`       | `boolean`           | `true`     | 构建编辑格式工具栏行，包括粗体、填充、插入和层叠顺序。该行仅在启用编辑时*可见*。                                                                               |
| `showInspector`           | `boolean`           | `true`     | 构建属性检查器面板，包括位置、尺寸、填充和线条。仅在启用编辑时*可见*。                                                                                         |
| `hiddenActions`           | `ToolbarActionId[]` | -          | 单独隐藏工具栏按钮或功能区选项卡，见下文。                                                                                                                     |
| `editable`                | `boolean`           | `false`    | 启用编辑，包括点击选择、拖动、缩放、旋转、内联文本编辑、键盘快捷键、撤销和重做，以及工具栏保存按钮。稍后可通过 [`setEditable`](/zh/vanilla/api#editing) 切换。 |
| `readOnly`                | `boolean`           | -          | 已由 `editable` 替代的旧标记，保留它是为了让已有选项对象继续通过类型检查，不再产生效果。                                                                       |
| `inspectorInitiallyOpen`  | `boolean`           | -          | 检查器的初始可见性；省略时保持桌面端或移动端的默认值。                                                                                                         |
| `rulerUnit`               | `RulerUnit`         | `'inches'` | “视图 > 标尺”刻度标签的单位制（`'centimetres'` 在各绑定中标签一致）。                                                                                          |

### `hiddenActions` {#hiddenactions}

`ToolbarActionId` 中每个 ID 控制一个快速访问按钮、一个控件组或整个功能区选项卡。与 `showToolbar` 不同，它可以隐藏单个部分，而不是整个界面：

- **按钮或控件组**：`'share'`、`'broadcast'`、`'export'`、`'undo'`、`'redo'`、`'record'`、`'notes'`、`'fullscreen'`、`'zoom'`（放大、缩小和适应作为一组）、`'navigation'`（上一张和下一张作为一组）。
- **功能区选项卡**：`'file'`、`'home'`、`'insert'`、`'draw'`、`'design'`、`'transitions'`、`'animations'`、`'slideShow'`、`'record'`、`'review'`、`'view'`、`'help'`。

`'record'` 同时隐藏快速访问录制控件和录制功能区选项卡，因为它们提供相同功能。

## 界面自定义 {#ui-customization}

| 选项            | 类型                  | 默认值 | 说明                                                                                                                                                                                                                                                         |
| --------------- | --------------------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `customization` | `ViewerCustomization` | -      | 隐藏、锁定或重新映射界面的任何部分：功能区选项卡和按钮、“文件 > 选项”的页面、分区和设置（锁定值、设置默认值）、文件选项卡的页面和卡片、右键菜单项、编辑器快捷键、面板、功能（AI、协作）和对话框。与 `hiddenActions` 取并集。传入新对象会替换整个自定义配置。 |

```ts
createPptxViewer(el, {
	source: bytes,
	editable: true,
	customization: { ribbon: { hiddenTabs: ['draw'] }, disabledFeatures: ['ai'] },
});
```

返回的实例上提供相同的辅助方法（`hideRibbonTab`、`lockSetting`、`remapShortcut` 等），可在运行时更改；参见[API 页面](/zh/vanilla/api#ui-customization)。全部 ID、规则和示例请参见[界面自定义指南](/zh/guide/customization)。

## 主题与本地化 {#theming--localization}

| 选项               | 类型                            | 默认值       | 说明                                                                                                                 |
| ------------------ | ------------------------------- | ------------ | -------------------------------------------------------------------------------------------------------------------- |
| `theme`            | `ViewerTheme`                   | -            | 查看器界面主题，共享的 `ViewerTheme` 包含颜色、圆角和 CSS 变量。参见[主题](/zh/vanilla/theming)。                    |
| `locale`           | `string`                        | `'en'`       | 界面语言，字典来自 `messages`，内置英文。                                                                            |
| `defaultLocale`    | `string`                        | -            | “文件 > 选项 > 语言”的初始语言。优先于已持久化的选择和 `locale`；用户之后的选择仍然生效。                            |
| `defaultThemeKey`  | `string`                        | -            | “文件 > 选项 > 外观”在 `availableThemes`（或内置目录）中的初始键。优先于已持久化的选择。                             |
| `onThemeChange`    | `(key: string) => void`         | -            | 主题选择的宿主回调；参见[回调](#callbacks)。                                                                         |
| `onLocaleChange`   | `(code: string) => void`        | -            | 语言选择的宿主回调；参见[回调](#callbacks)。                                                                         |
| `authorName`       | `string`                        | -            | 批注和广播演示者的作者名。回退到“选项 > 常规 > 用户名”，再回退到通用标签。                                           |
| `messages`         | `TranslationMessages`           | -            | 按语言提供的 `pptx.*` 消息字典。英文回退到内置共享字典，其他语言回退到英文。                                         |
| `availableThemes`  | `readonly ThemeCatalogEntry[]`  | 共享主题目录 | “文件 > 选项 > 外观”提供的主题选项，包括默认、亮色、朱红亮色和朱红深色，也会在设计选项卡的快速访问主题库中高亮。     |
| `availableLocales` | `readonly LocaleCatalogEntry[]` | 已注册的字典 | “文件 > 选项 > 语言”提供的语言选项，默认为注册了 `messages` 字典的所有语言，再加上 `'en'`。                          |
| `accountAuth`      | `AccountAuthConfig`             | 禁用         | 将真实登录流程接入“文件 > 账户”的可选入口（`{ enabled, onSignIn, signedInUser? }`），只有 `enabled: true` 时才渲染。 |

`TranslationMessages` 是 `Record<string, Record<string, string>>`：语言代码映射到以点分 `pptx.*` 键组成的扁平字典。稍后可以通过 [`setTheme` / `setLocale`](/zh/vanilla/api#theming--localization) 更改主题和语言。

`ThemeCatalogEntry` 是 `{ key: string; labelKey: string; theme: ViewerTheme | undefined }`，其中 `undefined` 恢复内置默认值；`LocaleCatalogEntry` 是 `{ code: string; label: string; nativeLabel: string }`。

## 扩展 {#extension}

| 选项                                                                       | 类型                        | 默认值                    | 说明                                                                                                                                                                                               |
| -------------------------------------------------------------------------- | --------------------------- | ------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `registry`                                                                 | `ElementRendererRegistry`   | `createDefaultRegistry()` | 自定义元素渲染器注册表。传入自己的注册表，或通过 `getRegistry()` 修改默认注册表，以添加或覆盖元素渲染器。参见[元素渲染器](/zh/vanilla/renderers)。                                                 |
| `smartArt3D`                                                               | `boolean`                   | `false`                   | 可选的 WebGL SmartArt 渲染器，将 `smartArt` 元素渲染为拉伸的 Three.js 场景。`three` 是可选 peer 依赖，仅当值为 `true` 时延迟导入；不可用时使用 SVG 渲染器。仅在构造时设置一次，没有运行时 setter。 |
| `surfaceChart3D`, `barChart3D`, `lineChart3D`, `areaChart3D`, `pieChart3D` | `boolean`                   | `false`                   | 分别启用对应三维图表类型的交互式 Three.js 渲染器。缺少 `three`，或图表无法渲染为 WebGL 场景时，均回退为 SVG。                                                                                      |
| `ai`                                                                       | `PptxAiConfig`              | -                         | 启用可选的 AI 助手，其 SDK peer 依赖仅在打开面板时加载。省略此选项则不提供助手。                                                                                                                   |
| `ribbonAddIns`                                                             | `readonly RibbonAddInTab[]` | -                         | 宿主在“帮助”之后添加的功能区选项卡。每个命令运行其 `run` 回调，并分发一个冒泡的 `office-ribbon-add-in` 事件。实例上的 `setRibbonAddIns` 可实时更改它们。                                           |

## 自动保存 {#autosave}

在共享 IndexedDB 存储中保存崩溃恢复快照，有未保存修改时每个间隔写入一次。自动保存不会替代用户真正的保存操作，只在下次启动时提供恢复保障。

| 选项                 | 类型      | 默认值                | 说明                                                                                       |
| -------------------- | --------- | --------------------- | ------------------------------------------------------------------------------------------ |
| `autosave`           | `boolean` | `true`                | 用于恢复的自动保存，工具栏显示小型状态标记。它是标题栏开关可启用范围的策略上限，详见下文。 |
| `autosaveIntervalMs` | `number`  | “文件 > 选项”中的间隔 | 有未保存修改时，两次恢复快照之间的间隔，单位为毫秒。显式值优先于用户的自动恢复设置。       |
| `autosaveFilePath`   | `string`  | `'presentation.pptx'` | 自动保存使用的 IndexedDB 恢复键。                                                          |

运行时控制由实例提供：[`autosaveNow` / `setAutosaveEnabled` / `isAutosaveEnabled`](/zh/vanilla/api#autosave)。

### autosave 属性和自动保存开关，谁优先 {#autosave-policy}

**五种绑定**遵循相同规则，统一实现在共享决策函数 `resolveAutosaveActivation` 中：

> **`autosave` 属性决定策略上限，标题栏开关表达该范围内的用户偏好。**

| `autosave` | 运行行为                                     | 开关                         |
| ---------- | -------------------------------------------- | ---------------------------- |
| 省略       | 允许自动保存，由用户开关决定，默认**开启**。 | 可用。                       |
| `true`     | 与省略相同，宿主允许，由用户决定。           | 可用。                       |
| `false`    | 自动保存关闭，加载时也不会提供恢复提示。     | **不可操作**，状态不能切换。 |

用户偏好不能越过宿主策略，因此 `autosave: false` 也会移除开关，避免出现看似可操作却没有效果的控件。无论采用哪种方式，`canEdit` / `editable` 和 `filePath` 键都是必要条件。

保存频率同样遵循该规则：显式 `autosaveIntervalMs` 是宿主策略，会按给定值执行；省略时遵循用户的 **“文件 > 选项 > 保存 > 每隔 N 分钟保存自动恢复信息”** 设置，默认两分钟。

默认值为 `true`，因为默认关闭的崩溃恢复无法为用户提供保障。

### 恢复快照 {#recovering-a-snapshot}

文稿加载完成后，如果同一键下存在 24 小时内的快照，查看器会弹出 **“恢复未保存的更改？”** 对话框，提供恢复和放弃选项。恢复会加载快照字节；放弃会删除快照。如果当前标签页已经接收过该快照，例如宿主已通过 `restoreSessionDeck` 恢复，则不会再次提示。

## 协作 {#collaboration}

| 选项            | 类型                 | 默认值 | 说明                                                                                                                                                                              |
| --------------- | -------------------- | ------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `collaboration` | `CollaborationInput` | -      | 立即启动实时协作会话，使用 Yjs 配合 y-websocket 或无服务器的 y-webrtc：可以是旧版 `CollaborationConfig`，也可以是共享的 `CollaborationOptions`。`role: 'viewer'` 配置会强制只读。 |
| `shareDefaults` | `ShareDefaults`      | -      | 内置共享和广播对话框表单字段的预填值（`{ roomId?, userName?, serverUrl? }`）；广播对话框使用 `userName` 作为演示者的显示名称。                                                    |

也可以稍后通过 [`startCollaboration` / `stopCollaboration`](/zh/vanilla/api#collaboration) 启动或停止会话。

::: warning 通信格式限制
媒体、OLE、三维模型和墨迹的二进制载荷不会通过网络传输，这是共享编解码器的限制。远程更新会替换整个本地幻灯片数组，因此加入者由宿主提供的媒体可能发生降级。
:::

## 自有 provider，或在同一浏览器内共享 {#your-own-provider-or-sharing-in-one-browser}

除 `CollaborationConfig` 外，`collaboration` 选项也接受 Word、Excel 和 Visio 共用的契约（来自 `ooxml-core/collab` 的 `CollaborationOptions`）。以下两种形式不需要宿主自管的 `session`；查看器自己打开会话，并在协作停止时销毁它：

```ts
createPptxViewer(el, {
	collaboration: { roomId: 'deck-review', provider, user: { name: 'Ada' }, intent: 'join' },
});
// Same browser only:
createPptxViewer(el, { collaboration: { roomId: 'deck-review', user: { name: 'Ada' } } });
```

- `provider`（`ProviderFactory`，例如 `transportProvider`，或通过 `adaptYjsProvider` 包装的现有 provider）优先于 `serverUrl`。
- 没有 `session`、`provider`、`serverUrl` 或 `transport` 时，房间通过 BroadcastChannel（`ooxml-pptx:<roomId>`）在此浏览器的标签页和窗口之间共享，不需要服务器。
- 优先级为 `session`（借用，从不销毁）> `provider` > `serverUrl` > BroadcastChannel。带 `serverUrl` 的旧版 `CollaborationConfig` 仍使用 y-websocket；带 `transport: 'webrtc'` 或空 `serverUrl` 的则仍使用无服务器的 y-webrtc 传输。
- `intent: 'join'` 永远不会用本地演示文稿初始化空房间。会话在同步完成后（或单个对等端的宽限期之后）变为可写；`role: 'viewer'` 保持只读。
- 会话打开期间，状态报告 `status: 'connecting'`；房间无效或 BroadcastChannel 不可用时报告 `error`。

## 回调 {#callbacks}

`PptxViewerOptions` 继承 `PptxViewerCallbacks`。这里没有框架事件系统，事件就是普通的回调选项：

| 回调                    | 签名                                                             | 说明                                                                                                                                                                                               |
| ----------------------- | ---------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `onLoad`                | `(info: { slideCount: number; canvasSize: CanvasSize }) => void` | 演示文稿成功加载后触发。                                                                                                                                                                           |
| `onError`               | `(message: string, error: unknown) => void`                      | 加载失败时触发，消息已经尽可能本地化。                                                                                                                                                             |
| `onSlideChange`         | `(index: number) => void`                                        | 当前幻灯片变化时触发，索引从 0 开始。                                                                                                                                                              |
| `onZoomChange`          | `(scale: number) => void`                                        | 实际缩放比例变化时触发，1 表示 100%。                                                                                                                                                              |
| `onPresentationChange`  | `(presenting: boolean) => void`                                  | 进入或退出全屏放映模式时触发。                                                                                                                                                                     |
| `onChange`              | `() => void`                                                     | 任何文档修改后触发，例如移动、缩放、编辑和撤销。                                                                                                                                                   |
| `onDirtyChange`         | `(dirty: boolean) => void`                                       | 未保存编辑标记变化时触发，保存会重置该标记。                                                                                                                                                       |
| `onSelectionChange`     | `(elementId: string \| null) => void`                            | 选中元素变化时触发，`null` 表示未选中。                                                                                                                                                            |
| `onAutosaveStatus`      | `(status: AutosaveStatus) => void`                               | 自动保存生命周期每次变化时触发（`'idle' \| 'saving' \| 'saved' \| 'error'`）。                                                                                                                     |
| `onAutosaveRecovery`    | `(record: AutosaveRecord) => void`                               | 提供启动时发现的恢复快照，由宿主决定是否恢复，见下文。                                                                                                                                             |
| `onCollaborationStatus` | `(status: ConnectionStatus) => void`                             | 协作连接状态每次变化时触发（`'disconnected' \| 'connecting' \| 'connected' \| 'error'`）。                                                                                                         |
| `onCollaborationChange` | `(state: CollaborationState) => void`                            | 会话启动或停止（关闭状态）时，以及状态、同步或人员变化时触发（也适用于宿主自管的会话），携带共享状态（`active`、`roomId`、`status`、`synced`、`people`）。`getCollaborationState()` 可按需返回它。 |
| `onThemeChange`         | `(key: string) => void`                                          | 通过“文件 > 选项 > 外观”选择主题，或调用 `setTheme` 且匹配目录条目时触发。提供回调后由宿主管理持久化，否则查看器使用 `localStorage` 中的 `pptx-viewer-prefs`。                                     |
| `onLocaleChange`        | `(code: string) => void`                                         | 通过“文件 > 选项 > 语言”选择语言，或任何 `setLocale` 调用时触发。持久化规则与 `onThemeChange` 相同。                                                                                               |
| `onToggleAutosave`      | `(enabled: boolean) => void`                                     | 标题栏自动保存控件启用或禁用恢复自动保存时触发。                                                                                                                                                   |
| `onContentChange`       | `(content: Uint8Array) => void`                                  | 每次提交编辑后，携带序列化的 `.pptx` 字节触发。序列化是异步的，因此连续的多次编辑可能按顺序交付多个快照。                                                                                          |
| `onModeChange`          | `(mode: ViewerMode) => void`                                     | 查看器模式变化时触发（`'preview'`、`'edit'`、`'present'`、`'master'`）。                                                                                                                           |
| `onSlideCountChange`    | `(count: number) => void`                                        | 幻灯片总数变化时触发（添加、删除幻灯片或加载了演示文稿）。                                                                                                                                         |
| `onStartCollaboration`  | `(config: CollaborationConfig) => void`                          | 用户从共享或广播对话框启动会话时触发。                                                                                                                                                             |
| `onStopCollaboration`   | `() => void`                                                     | 用户从共享或广播对话框停止会话时触发。                                                                                                                                                             |
| `onOpenFile`            | `() => void`                                                     | 宿主对“文件 > 打开”的覆盖：绕过内置选择器，由宿主通过 `loadFile` 提供演示文稿。                                                                                                                    |
| `onHyperlinkClick`      | `HyperlinkClickHandler`                                          | 跟随链接之前触发；返回 `false` 可取消默认操作。                                                                                                                                                    |

`AutosaveRecord` 是 `{ key: string; data: Uint8Array; timestamp: number; size: number }`，典型恢复流程为 `viewer.loadFile(record.data)`。

## 完整接口 {#full-interface}

```ts
interface PptxViewerCallbacks {
	onLoad?: (info: { slideCount: number; canvasSize: CanvasSize }) => void;
	onError?: (message: string, error: unknown) => void;
	onSlideChange?: (index: number) => void;
	onZoomChange?: (scale: number) => void;
	onPresentationChange?: (presenting: boolean) => void;
	onChange?: () => void;
	onDirtyChange?: (dirty: boolean) => void;
	onSelectionChange?: (elementId: string | null) => void;
	onAutosaveStatus?: (status: AutosaveStatus) => void;
	onAutosaveRecovery?: (record: AutosaveRecord) => void;
	onCollaborationStatus?: (status: ConnectionStatus) => void;
	onCollaborationChange?: (state: CollaborationState) => void;
	onContentChange?: (content: Uint8Array) => void;
	onModeChange?: (mode: ViewerMode) => void;
	onSlideCountChange?: (count: number) => void;
	onStartCollaboration?: (config: CollaborationConfig) => void;
	onStopCollaboration?: () => void;
	onOpenFile?: () => void;
	onHyperlinkClick?: HyperlinkClickHandler;
}

interface PptxViewerOptions extends PptxViewerCallbacks {
	source?: PptxViewerSource;
	fonts?: ViewerFontSource[];
	theme?: ViewerTheme;
	fileName?: string;
	locale?: string;
	defaultLocale?: string;
	defaultThemeKey?: string;
	authorName?: string;
	messages?: TranslationMessages;
	availableThemes?: readonly ThemeCatalogEntry[];
	availableLocales?: readonly LocaleCatalogEntry[];
	onThemeChange?: (key: string) => void;
	onLocaleChange?: (code: string) => void;
	accountAuth?: AccountAuthConfig;
	initialSlide?: number;
	editable?: boolean;
	readOnly?: boolean;
	showToolbar?: boolean;
	showThumbnails?: boolean;
	showCompatibilityToasts?: boolean;
	showFormatToolbar?: boolean;
	showInspector?: boolean;
	inspectorInitiallyOpen?: boolean;
	rulerUnit?: RulerUnit;
	customization?: ViewerCustomization;
	ribbonAddIns?: readonly RibbonAddInTab[];
	hiddenActions?: ToolbarActionId[];
	registry?: ElementRendererRegistry;
	smartArt3D?: boolean;
	surfaceChart3D?: boolean;
	barChart3D?: boolean;
	lineChart3D?: boolean;
	areaChart3D?: boolean;
	pieChart3D?: boolean;
	ai?: PptxAiConfig;
	autosave?: boolean;
	onToggleAutosave?: (enabled: boolean) => void;
	autosaveIntervalMs?: number;
	autosaveFilePath?: string;
	collaboration?: CollaborationInput;
	shareDefaults?: ShareDefaults;
}
```

## 示例：完整接线 {#example-everything-wired}

```ts
import { createPptxViewer, vermilionLightTheme } from 'pptx-vanilla-viewer';

const viewer = createPptxViewer(document.getElementById('host')!, {
	source: '/decks/quarterly.pptx',
	fileName: 'quarterly.pptx',
	theme: vermilionLightTheme,
	locale: 'en',
	initialSlide: 0,
	editable: true,
	showToolbar: true,
	showThumbnails: true,
	hiddenActions: ['broadcast', 'record'],
	autosave: true,
	autosaveFilePath: 'quarterly.pptx',
	onLoad: ({ slideCount }) => console.log(`${slideCount} slides`),
	onSlideChange: (index) => console.log('slide', index + 1),
	onZoomChange: (scale) => console.log(`${Math.round(scale * 100)}%`),
	onPresentationChange: (presenting) => console.log(presenting ? 'presenting' : 'back'),
	onDirtyChange: (dirty) => console.log('unsaved edits:', dirty),
	onSelectionChange: (elementId) => console.log('selected', elementId),
	onAutosaveStatus: (status) => console.log('autosave:', status),
	onAutosaveRecovery: (record) => {
		if (confirm('Restore unsaved changes from your last session?')) {
			void viewer.loadFile(record.data);
		}
	},
	onError: (message) => console.error(message),
});
```
