---
title: 组件输入与输出
description: PowerPointViewerComponent 的 @Input() 和 @Output() 完整参考，涵盖内容、编辑、事件、主题和协作。
---

# 组件输入与输出 {#component-inputs-outputs}

自定义宿主尺寸时，请参阅[视口适配](/zh/guide/viewport-fit)，了解 `fitPadding`、`maxFitScale`、各框架示例和默认值。

`PowerPointViewerComponent`（选择器为 `pptx-viewer`）通过基于信号的 `input()` 提供配置，通过 `output()` 发出事件。只读查看器只需要 `content`，其他配置均可选。此参考直接依据组件源码整理。

```ts
import { PowerPointViewerComponent } from 'pptx-angular-viewer';
```

::: tip 提示
Angular 没有 `forwardRef` 句柄。`PowerPointViewerComponent` 的公开方法，包括导航、撤销与重做、缩放、`getContent()` 等，直接位于组件实例上，可以通过模板引用变量或 `viewChild()` 访问。参见[公开 API](/zh/angular/api)。
:::

## 内容 {#content}

| 输入       | 类型                                | 默认值 | 说明                                                                        |
| ---------- | ----------------------------------- | ------ | --------------------------------------------------------------------------- |
| `content`  | `Uint8Array \| ArrayBuffer \| null` | `null` | 原始 `.pptx` 文件字节。可直接接受类型化数组或 `ArrayBuffer`，无需手动包装。 |
| `filePath` | `string \| undefined`               | -      | 原始文件路径或名称，用作 IndexedDB 自动保存恢复的键，并显示在标题栏中。     |
| `fileName` | `string \| undefined`               | -      | 显示在标题栏保存位置状态旁的名称，回退为本地化的“演示文稿”。                |

::: tip 自动保存恢复需要 `filePath`
查看器内置的自动保存定时器会定期将文档序列化到 IndexedDB，以 `filePath` 为键。重新加载后若没有稳定的 `filePath`，就没有可供查找的恢复键。
:::

## 自动保存 {#autosave-policy}

| 输入                 | 类型                   | 默认值                | 说明                                                             |
| -------------------- | ---------------------- | --------------------- | ---------------------------------------------------------------- |
| `autosave`           | `boolean \| undefined` | `true`                | 用于恢复的自动保存，是标题栏开关可启用范围的策略上限，详见下文。 |
| `autosaveIntervalMs` | `number \| undefined`  | “文件 > 选项”中的间隔 | 恢复快照的保存间隔。显式值优先于用户的自动恢复设置。             |

**五种绑定**遵循相同规则，统一实现在共享决策函数 `resolveAutosaveActivation` 中：

> **`autosave` 输入是宿主策略上限，标题栏的自动保存开关是用户在此范围内的偏好。**

| `autosave` | 运行行为                                     | 开关                         |
| ---------- | -------------------------------------------- | ---------------------------- |
| 省略       | 允许自动保存，由用户开关决定，默认**开启**。 | 可用。                       |
| `true`     | 与省略相同，宿主允许，由用户决定。           | 可用。                       |
| `false`    | 自动保存关闭，加载时也不会提供恢复提示。     | **不可操作**，状态不能切换。 |

用户偏好不能越过宿主策略，因此 `autosave="false"` 也会移除开关。无论采用哪种方式，`canEdit` 和 `filePath` 键都是必要条件。显式 `autosaveIntervalMs` 优先于用户在 **“文件 > 选项 > 保存 > 每隔 N 分钟保存自动恢复信息”** 中设置的频率，默认两分钟。

文稿加载完成后，如果同一 `filePath` 下存在 24 小时内的快照，查看器会弹出 **“恢复未保存的更改？”** 对话框，提供恢复和放弃选项。

## 编辑 {#editing}

| 输入      | 类型      | 默认值  | 说明                                                                                       |
| --------- | --------- | ------- | ------------------------------------------------------------------------------------------ |
| `canEdit` | `boolean` | `false` | 启用编辑操作，包括功能区编辑工具、检查器编辑、内联文本编辑和幻灯片管理。`false` 表示只读。 |

## 显示 {#display}

这些属性与 Svelte 和原生绑定一致，默认值和范围限制由共享代码统一处理。

| 输入                        | 类型                           | 默认值 | 说明                                                                                                                                                                                                                                                                                                                                                                                                          |
| --------------------------- | ------------------------------ | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `[initialSlide]`            | `number`                       | `0`    | 每次加载后显示的幻灯片，索引从 0 开始，自动限制在有效范围内。之后修改只对下一次加载生效。                                                                                                                                                                                                                                                                                                                     |
| `[googleFonts]`             | `false \| { baseUrl: string }` | 未设置 | 未安装也未嵌入的字体的 Google Fonts 回退。未设置时从 `fonts.googleapis.com` 加载；`false` 不向 Google 发送任何请求（文字使用已安装、嵌入或回退字体）；`{ baseUrl }` 改为向自托管镜像发送 CSS2 请求。参见[字体与隐私](../guide/runtime-environments.md#fonts-and-privacy)。                                                                                                                                    |
| `[allowExternalContent]`    | `boolean`                      | 未设置 | 演示文稿从其他服务器链接的内容（http(s) 目标）的宿主开关：图片、图片填充和项目符号、视频、音频和海报帧。`false` 不获取任何此类内容，`true` 允许获取；两者都会将 文件 > 选项 > 信任中心 >“允许外部内容”锁定为对应值。未设置时由用户在该处的选择决定。更改从下一次加载起生效。Google Fonts 请求由 `googleFonts` 单独控制。参见[外部内容与隐私](../guide/runtime-environments.md#external-content-and-privacy)。 | 未安装也未嵌入的字体的 Google Fonts 回退。未设置时从 `fonts.googleapis.com` 加载；`false` 不向 Google 发送任何请求（文字使用已安装、嵌入或回退字体）；`{ baseUrl }` 改为向自托管镜像发送 CSS2 请求。参见[字体与隐私](../guide/runtime-environments.md#fonts-and-privacy)。 |
| `[limits]`                  | `PptxLoadLimits`               | 未设置 | 加载演示文稿的工作量上限，供显示不受信任演示文稿的宿主使用：`maxInputBytes`、`maxUncompressedBytes`、`maxEntries`、`maxSlides`、`maxShapes`。未设置时保留核心默认值（解压后 500 MiB、65,536 个 ZIP 条目，其他不限）。超出限制的演示文稿会显示加载错误界面，并给出限制值和演示文稿的实际值。更改从下一次加载起生效。参见[加载限制](../guide/runtime-environments.md#load-limits)。                             |
| `[showToolbar]`             | `boolean`                      | `true` | 显示编辑器外框：标题栏、功能区和工具栏、受保护视图与只读横幅、移动端工具栏和状态栏。为 `false` 时只保留画布（以及显示中的缩略图窗格）。                                                                                                                                                                                                                                                                       |
| `[showThumbnails]`          | `boolean`                      | `true` | 显示幻灯片缩略图窗格。仍受用户自己的折叠开关和 `slidesPane` 自定义面板控制。                                                                                                                                                                                                                                                                                                                                  |
| `[showCompatibilityToasts]` | `boolean`                      | `true` | 加载后显示兼容性警告提示堆栈。`false` 只隐藏堆栈，警告仍会被收集。不受 `showToolbar` 影响。                                                                                                                                                                                                                                                                                                                   |

```html
<pptx-viewer [content]="bytes" [initialSlide]="2" [showToolbar]="false" [showThumbnails]="false" />
```

## 事件（`@Output()`） {#events-output-s}

| 输出                  | 载荷                          | 说明                                                                                        |
| --------------------- | ----------------------------- | ------------------------------------------------------------------------------------------- |
| `dirtyChange`         | `boolean`                     | 未保存修改标记切换时触发。                                                                  |
| `contentChange`       | `Uint8Array`                  | 每当 [`getContent()`](/zh/angular/api) 将文稿序列化为字节时触发，携带重新序列化的文档字节。 |
| `activeSlideChange`   | `number`                      | 当前幻灯片变化时触发。                                                                      |
| `modeChange`          | `string`                      | 查看器模式变化时触发，模式为 `'preview'`、`'edit'`、`'present'` 或 `'master'`。             |
| `zoomChange`          | `number`                      | 缩放级别变化时触发。                                                                        |
| `selectionChange`     | `string[]`                    | 元素选区变化时触发。                                                                        |
| `slideCountChange`    | `number`                      | 幻灯片总数变化时触发，例如添加或删除幻灯片。                                                |
| `propertiesChange`    | `Partial<PptxCoreProperties>` | 用户在信息对话框中编辑文档属性（标题、作者等）时触发。                                      |
| `startCollaboration`  | `CollaborationConfig`         | 从共享或广播对话框启动会话时触发，请据此设置 `collaboration` 输入。                         |
| `stopCollaboration`   | `void`                        | 从共享或广播对话框停止会话时触发，请据此清空 `collaboration` 输入。                         |
| `collaborationChange` | `CollaborationState`          | 会话启动或停止，以及其状态、已同步标志或人员变化时触发。                                    |
| `load`                | `ViewerLoadDetail`            | 每次演示文稿无错误地加载完成时触发（`slideCount`、`canvasSize`）。                          |
| `error`               | `string`                      | 每种不同的加载失败触发一次，携带易读的消息。                                                |

::: info 说明
`contentChange` 提供的是 `Uint8Array`，不是布尔值。需要按需拉取内容时，请调用组件实例上的 [`getContent()`](/zh/angular/api)。
:::

## 展示与创作 {#presentation-authoring}

| 输入                                                                       | 类型                                     | 默认值  | 说明                                                                                                                                                                |
| -------------------------------------------------------------------------- | ---------------------------------------- | ------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `authorName`                                                               | `string \| undefined`                    | -       | 用于批注和标注作者、以及广播所有者的显示名称，回退为 `collaboration.userName` 或 `'You'` / `'Presenter'`。                                                          |
| `class`                                                                    | `string`                                 | `''`    | 应用于查看器根元素的可选类名。                                                                                                                                      |
| `smartArt3D`                                                               | `boolean`                                | `false` | 启用基于 WebGL 的 Three.js SmartArt 挤出三维块渲染器，需要可选同级依赖 `three`，缺少时回退到 SVG。                                                                  |
| `surfaceChart3D`, `barChart3D`, `lineChart3D`, `areaChart3D`, `pieChart3D` | `boolean`                                | `false` | 分别启用对应三维图表类型的交互式 Three.js 渲染器。缺少 `three`，或图表无法渲染为 WebGL 场景时，均回退为 SVG。                                                       |
| `ai`                                                                       | `PptxAiAssistant \| PptxAiConfig`        | -       | 启用可选智能助手：传入 `pptx-angular-viewer/ai` 中的 `createAiAssistant(config)`，或普通的 `PptxAiConfig`。SDK 同级依赖只在打开面板时加载，省略此输入则不提供助手。 |
| `onOpenFile`                                                               | `(() => void) \| undefined`              | -       | 由宿主接管“文件 > 打开”操作，绕过内置原生文件选择器，随后由宿主提供新的 `content` 值。                                                                              |
| `onHyperlinkClick`                                                         | `HyperlinkClickHandler`                  | -       | 函数输入（不是输出）。跟随链接之前调用；返回 `false` 可取消默认操作。                                                                                               |
| `ribbonAddIns`                                                             | `readonly RibbonAddInTab[] \| undefined` | -       | 宿主在“帮助”之后添加的功能区选项卡。每个命令会运行其 `run` 回调，并派发冒泡的 `office-ribbon-add-in` 事件。仅限桌面功能区。                                         |
| `inspectorInitiallyOpen`                                                   | `boolean \| undefined`                   | -       | 检查器的初始可见性；省略时保持桌面端或移动端的默认值。                                                                                                              |
| `fonts`                                                                    | `ViewerFontSource[]`                     | `[]`    | 由宿主提供的授权字体来源。不内置任何字体。                                                                                                                          |
| `hiddenActions`                                                            | `ToolbarActionId[]`                      | `[]`    | 要隐藏的单个工具栏按钮和（或）功能区选项卡。与 `customization` 取并集。                                                                                             |

## 主题 {#theming}

| 输入    | 类型                       | 默认值 | 说明                                                                                                                          |
| ------- | -------------------------- | ------ | ----------------------------------------------------------------------------------------------------------------------------- |
| `theme` | `ViewerTheme \| undefined` | -      | 主题配置，包括部分颜色覆盖、自定义 `radius` 和任意 `cssVars`。未设置的值回退为内置深色主题。参见[主题](/zh/angular/theming)。 |

```ts
<pptx-viewer
  [content]="bytes"
  [theme]="{ colors: { primary: '#6366f1', background: '#0f172a' }, radius: '0.75rem' }"
/>
```

### 主题、语言和账户 {#theme-language-and-account}

| 输入               | 类型                                         | 说明                                                       |
| ------------------ | -------------------------------------------- | ---------------------------------------------------------- |
| `defaultThemeKey`  | `string \| undefined`                        | “文件 > 选项 > 主题”的初始主题键。                         |
| `availableThemes`  | `readonly ThemeCatalogEntry[] \| undefined`  | “文件 > 选项”提供的主题选项。                              |
| `onThemeChange`    | `((key: string) => void) \| undefined`       | 函数输入。用户选择主题时调用；之后由宿主负责持久化该选择。 |
| `defaultLocale`    | `string \| undefined`                        | “文件 > 选项 > 语言”的初始语言代码。                       |
| `availableLocales` | `readonly LocaleCatalogEntry[] \| undefined` | “文件 > 选项 > 语言”提供的语言选项。                       |
| `onLocaleChange`   | `((code: string) => void) \| undefined`      | 函数输入。用户选择语言时调用；之后由宿主负责应用并持久化。 |
| `accountAuth`      | `AccountAuthConfig \| undefined`             | “文件 > 账户”登录部分的接入点；默认禁用。                  |

## 界面定制 {#ui-customization}

| 输入              | 类型                  | 默认值 | 说明                                                                                                                                                                                                                                                   |
| ----------------- | --------------------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `[customization]` | `ViewerCustomization` | -      | 隐藏、锁定或重新映射界面的任何部分：功能区选项卡和按钮、“文件 > 选项”的页面、分区和设置（锁定值、设置默认值）、文件选项卡的页面和卡片、右键菜单项、编辑器快捷键、面板、功能（AI、协作）和对话框。与 `hiddenActions` 取并集。传入新对象会替换整个定制。 |

```html
<pptx-viewer [content]="bytes" [canEdit]="true" [customization]="customization" />
```

同样的辅助方法（`hideRibbonTab`、`lockSetting`、`remapShortcut` 等）也可在组件实例上调用，用于在运行时修改定制，参见[API 页面](/zh/angular/api#ui-customization)。所有 ID、规则和用法示例见[界面定制指南](/zh/guide/customization)。

## 协作 {#collaboration}

这些输入用于启用和控制实时协同编辑。完整流程及 `CollaborationConfig` 结构请参见[实时协作](/zh/angular/collaboration)。

| 输入            | 类型                                                                      | 默认值 | 说明                                                                                                                                                                                 |
| --------------- | ------------------------------------------------------------------------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `collaboration` | `CollaborationInput \| undefined`                                         | -      | 提供后启用协同编辑，包括实时光标、在线状态和 Yjs CRDT 同步。接受旧版 `CollaborationConfig` 或共享的 `CollaborationOptions`。需要 `yjs` peer 依赖，以及 `y-websocket` 或 `y-webrtc`。 |
| `shareDefaults` | `{ roomId?: string; userName?: string; serverUrl?: string } \| undefined` | -      | 共享或广播对话框字段的默认值。省略时字段为空，userName 回退为 `authorName`。                                                                                                         |

::: info 说明
与 React 和 Vue 一样，`collaboration` 是受控的。查看器不会自行启动会话，请监听 `startCollaboration` 以设置输入，监听 `stopCollaboration` 以清空输入。
:::

## 完整输入与输出列表 {#full-input-output-list}

```ts
class PowerPointViewerComponent {
	// Inputs
	readonly content = input<Uint8Array | ArrayBuffer | null>(null);
	readonly canEdit = input<boolean>(false);
	readonly initialSlide = input<number | undefined>(undefined);
	readonly googleFonts = input<false | { baseUrl: string } | undefined>(undefined);
	readonly allowExternalContent = input<boolean | undefined>(undefined);
	readonly limits = input<PptxLoadLimits | undefined>(undefined);
	readonly showToolbar = input<boolean | undefined>(undefined);
	readonly showThumbnails = input<boolean | undefined>(undefined);
	readonly showCompatibilityToasts = input<boolean | undefined>(undefined);
	readonly class = input<string>('');
	readonly theme = input<ViewerTheme | undefined>(undefined);
	readonly filePath = input<string | undefined>(undefined);
	readonly fileName = input<string | undefined>(undefined);
	readonly collaboration = input<CollaborationInput | undefined>(undefined);
	readonly authorName = input<string>();
	readonly shareDefaults = input<
		{ roomId?: string; userName?: string; serverUrl?: string } | undefined
	>(undefined);
	readonly onOpenFile = input<(() => void) | undefined>(undefined);
	readonly smartArt3D = input<boolean>(false);
	readonly surfaceChart3D = input<boolean>(false);
	readonly barChart3D = input<boolean>(false);
	readonly lineChart3D = input<boolean>(false);
	readonly areaChart3D = input<boolean>(false);
	readonly pieChart3D = input<boolean>(false);
	readonly ai = input<PptxAiConfig | undefined>(undefined);
	readonly onHyperlinkClick = input<HyperlinkClickHandler>();
	readonly ribbonAddIns = input<readonly RibbonAddInTab[] | undefined>(undefined);
	readonly inspectorInitiallyOpen = input<boolean | undefined>(undefined);
	readonly defaultThemeKey = input<string | undefined>(undefined);
	readonly availableThemes = input<readonly ThemeCatalogEntry[] | undefined>(undefined);
	readonly onThemeChange = input<((key: string) => void) | undefined>(undefined);
	readonly defaultLocale = input<string | undefined>(undefined);
	readonly availableLocales = input<readonly LocaleCatalogEntry[] | undefined>(undefined);
	readonly onLocaleChange = input<((code: string) => void) | undefined>(undefined);
	readonly accountAuth = input<AccountAuthConfig | undefined>(undefined);

	// Outputs
	readonly activeSlideChange = output<number>();
	readonly dirtyChange = output<boolean>();
	readonly contentChange = output<Uint8Array>();
	readonly propertiesChange = output<Partial<PptxCoreProperties>>();
	readonly modeChange = output<string>();
	readonly zoomChange = output<number>();
	readonly selectionChange = output<string[]>();
	readonly slideCountChange = output<number>();
	readonly startCollaboration = output<CollaborationConfig>();
	readonly stopCollaboration = output<void>();
	readonly collaborationChange = output<CollaborationState>();
	readonly load = output<ViewerLoadDetail>();
	readonly error = output<string>();
}
```

## 功能触发方式 {#notes-on-triggering-features}

导航、模式切换、缩放、撤销与重做、选择操作均可以通过[公开 API](/zh/angular/api) 使用，例如 `viewer.goNext()`、`viewer.setMode('present')`。导出和打印通过内置功能区或对话框驱动，文档导出细节请参见[导出](/zh/angular/export)。
