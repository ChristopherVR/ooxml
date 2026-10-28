---
title: 组件属性
description: PowerPointViewerProps 参考，涵盖内容、编辑、回调、主题、协作以及自动保存与恢复。
---

# 组件属性 {#component-props}

自定义宿主尺寸时，请参阅[视口适配](/zh/guide/viewport-fit)，了解 `fitPadding`、`maxFitScale`、各框架示例和默认值。

`PowerPointViewer` 接收下文的 `PowerPointViewerProps` 接口。只有 `content` 必填，其余均可选。此参考直接对应源码接口。

```tsx
import { PowerPointViewer } from 'pptx-react-viewer';
import type { PowerPointViewerProps } from 'pptx-react-viewer';
```

::: tip 提示
`PowerPointViewer` 使用 `forwardRef`，因此还可传入 [`PowerPointViewerHandle`](/zh/react/handle) 类型的 `ref`，它不属于 `PowerPointViewerProps`。
:::

## 内容 {#content}

| 属性                 | 类型         | 默认值                | 说明                                                                                           |
| -------------------- | ------------ | --------------------- | ---------------------------------------------------------------------------------------------- |
| `content`            | `Uint8Array` | 必填                  | 原始 `.pptx` 字节。`ArrayBuffer` 需通过 `new Uint8Array(buf)` 包装。                           |
| `filePath`           | `string`     | -                     | 原始文件路径或名称，用作自动保存恢复的键，并显示在标题栏。                                     |
| `fileName`           | `string`     | 通用标签              | 打开文档的显示名称，显示在标题栏，例如 `quarterly-review.pptx`。                               |
| `autosave`           | `boolean`    | `true`                | 控制恢复快照的自动保存，是标题栏开关之上的宿主策略上限，详见[自动保存策略](#autosave-policy)。 |
| `autosaveIntervalMs` | `number`     | “文件 > 选项”中的间隔 | 恢复快照的保存间隔。显式值优先于用户的自动恢复设置。                                           |

::: warning content 类型
该属性使用 `Uint8Array`，不是 `ArrayBuffer`。从 `fetch` 或文件得到 `ArrayBuffer` 后的转换方式见[快速接入](/zh/react/getting-started)。
:::

::: tip 自动保存恢复需要 filePath
组件定时将文档序列化到 IndexedDB，以 `filePath` 为键。重载页面时，会检查相同 `filePath` 对应的近期快照。如果宿主没有持久保存并重新传入该值，恢复流程不会触发。详见下文[自动保存与恢复](#autosave-recovery)。
:::

## 编辑 {#editing}

| 属性      | 类型      | 默认值  | 说明                                                                        |
| --------- | --------- | ------- | --------------------------------------------------------------------------- |
| `canEdit` | `boolean` | `false` | 启用工具栏编辑、属性面板编辑、行内文本编辑和幻灯片管理。为 `false` 时只读。 |

## 显示 {#display}

这些属性与 Svelte 和原生绑定一致，默认值和范围限制由共享代码统一处理。

| 属性                      | 类型                           | 默认值 | 说明                                                                                                                                                                                                                                                                                                                                                                                                          |
| ------------------------- | ------------------------------ | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `initialSlide`            | `number`                       | `0`    | 每次加载后显示的幻灯片，索引从 0 开始，自动限制在有效范围内。之后修改只对下一次加载生效。                                                                                                                                                                                                                                                                                                                     |
| `googleFonts`             | `false \| { baseUrl: string }` | 未设置 | 未安装也未嵌入的字体的 Google Fonts 回退。未设置时从 `fonts.googleapis.com` 加载；`false` 不向 Google 发送任何请求（文字使用已安装、嵌入或回退字体）；`{ baseUrl }` 改为向自托管镜像发送 CSS2 请求。参见[字体与隐私](../guide/runtime-environments.md#fonts-and-privacy)。                                                                                                                                    |
| `allowExternalContent`    | `boolean`                      | 未设置 | 演示文稿从其他服务器链接的内容（http(s) 目标）的宿主开关：图片、图片填充和项目符号、视频、音频和海报帧。`false` 不获取任何此类内容，`true` 允许获取；两者都会将 文件 > 选项 > 信任中心 >“允许外部内容”锁定为对应值。未设置时由用户在该处的选择决定。更改从下一次加载起生效。Google Fonts 请求由 `googleFonts` 单独控制。参见[外部内容与隐私](../guide/runtime-environments.md#external-content-and-privacy)。 | 未安装也未嵌入的字体的 Google Fonts 回退。未设置时从 `fonts.googleapis.com` 加载；`false` 不向 Google 发送任何请求（文字使用已安装、嵌入或回退字体）；`{ baseUrl }` 改为向自托管镜像发送 CSS2 请求。参见[字体与隐私](../guide/runtime-environments.md#fonts-and-privacy)。 |
| `limits`                  | `PptxLoadLimits`               | 未设置 | 加载演示文稿的工作量上限，供显示不受信任演示文稿的宿主使用：`maxInputBytes`、`maxUncompressedBytes`、`maxEntries`、`maxSlides`、`maxShapes`。未设置时保留核心默认值（解压后 500 MiB、65,536 个 ZIP 条目，其他不限）。超出限制的演示文稿会显示加载错误界面，并给出限制值和演示文稿的实际值。更改从下一次加载起生效。参见[加载限制](../guide/runtime-environments.md#load-limits)。                             |
| `showToolbar`             | `boolean`                      | `true` | 显示编辑器外框：标题栏、功能区和工具栏、受保护视图与只读横幅、移动端工具栏和状态栏。为 `false` 时只保留画布（以及显示中的缩略图窗格）。                                                                                                                                                                                                                                                                       |
| `showThumbnails`          | `boolean`                      | `true` | 显示幻灯片缩略图窗格。仍受用户自己的折叠开关和 `slidesPane` 自定义面板控制。                                                                                                                                                                                                                                                                                                                                  |
| `showCompatibilityToasts` | `boolean`                      | `true` | 加载后显示兼容性警告提示堆栈。`false` 只隐藏堆栈，警告仍会被收集。不受 `showToolbar` 影响。                                                                                                                                                                                                                                                                                                                   |

```tsx
<PowerPointViewer content={bytes} initialSlide={2} showToolbar={false} showThumbnails={false} />
```

## 回调 {#callbacks}

| 属性                  | 类型                             | 默认值 | 说明                                                              |
| --------------------- | -------------------------------- | ------ | ----------------------------------------------------------------- |
| `onDirtyChange`       | `(isDirty: boolean) => void`     | -      | 未保存修改标记变化时调用。                                        |
| `onContentChange`     | `(content: Uint8Array) => void`  | -      | 内容变化时返回重新序列化后的文档字节。                            |
| `onActiveSlideChange` | `(slideIndex: number) => void`   | -      | 当前幻灯片变化时调用。                                            |
| `onModeChange`        | `(mode: ViewerMode) => void`     | -      | 模式变化时调用，例如从编辑切换到放映。                            |
| `onZoomChange`        | `(zoom: number) => void`         | -      | 缩放比例变化时调用。                                              |
| `onSelectionChange`   | `(elementIds: string[]) => void` | -      | 元素选择变化时调用。                                              |
| `onSlideCountChange`  | `(count: number) => void`        | -      | 添加或删除幻灯片导致总数变化时调用。                              |
| `onOpenFile`          | `() => void`                     | -      | 覆盖“文件 > 打开”，跳过内置选择器，由宿主随后传入新的 `content`。 |

::: info 说明
`onContentChange` 返回序列化文档的 `Uint8Array`，不是布尔值。需要主动获取内容时，使用句柄的 [`getContent()`](/zh/react/handle)。
:::

## 展示与创作 {#presentation-authoring}

| 属性                                                                       | 类型                              | 默认值  | 说明                                                                                                                                                        |
| -------------------------------------------------------------------------- | --------------------------------- | ------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `authorName`                                                               | `string`                          | -       | 批注和标记的作者显示名称。协作时回退到 `collaboration.userName`，否则回退到 `'You'`。                                                                       |
| `className`                                                                | `string`                          | -       | 添加到组件根元素的可选类名。                                                                                                                                |
| `smartArt3D`                                                               | `boolean`                         | `false` | 启用基于 WebGL 的 Three.js SmartArt 挤出三维块渲染器，需要可选同级依赖 `three`，缺少时回退到 SVG。                                                          |
| `surfaceChart3D`、`barChart3D`、`lineChart3D`、`areaChart3D`、`pieChart3D` | `boolean`                         | `false` | 分别启用对应三维图表的交互式 Three.js 渲染器。缺少 `three` 或图表无法作为 WebGL 场景渲染时，回退到对应 SVG 渲染器。                                         |
| `ai`                                                                       | `PptxAiAssistant \| PptxAiConfig` | -       | 启用可选智能助手：传入 `pptx-react-viewer/ai` 中的 `createAiAssistant(config)`，或普通的 `PptxAiConfig`。SDK 同级依赖只在打开面板时加载，省略则不提供助手。 |

## 主题 {#theming}

| 属性    | 类型          | 默认值 | 说明                                                                                                                |
| ------- | ------------- | ------ | ------------------------------------------------------------------------------------------------------------------- |
| `theme` | `ViewerTheme` | -      | 支持局部颜色覆盖、自定义 `radius` 和任意 `cssVars`。未设置项回退到内置深色主题，详见[主题配置](/zh/react/theming)。 |

```tsx
<PowerPointViewer
	content={bytes}
	theme={{
		colors: { primary: '#6366f1', background: '#0f172a' },
		radius: '0.75rem',
	}}
/>
```

## 语言 {#language}

| 属性               | 类型                     | 默认值                           | 说明                                                                                                                                                                                    |
| ------------------ | ------------------------ | -------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `defaultLocale`    | `string`                 | 已保存的选择，否则为 `'en'`      | “文件 > 选项 > 语言”的初始语言代码。依次回退到 `localStorage` 中保存的选择和 `'en'`。在 `PptxI18nProvider` 内忽略此项，语言由其 `locale` 决定。                                         |
| `availableLocales` | `LocaleCatalogEntry[]`   | `react-i18next` 中注册的语言代码 | “文件 > 选项 > 语言”提供的语言。默认使用宿主 `react-i18next` 实例中注册的语言代码。在 `PptxI18nProvider` 内没有默认值：需要同时传入至少两个语言和 `onLocaleChange` 才会显示“语言”类别。 |
| `onLocaleChange`   | `(code: string) => void` | -                                | 用户在“文件 > 选项 > 语言”中选择语言时调用。提供后由宿主负责应用和持久化；否则由组件自行应用和持久化。在 `PptxI18nProvider` 内语言始终由宿主掌控：未提供此回调时不显示“语言”类别。      |

如果要用自己的 i18n 库代替 react-i18next 翻译组件，请用 `PptxI18nProvider` 包裹它（仅限 React），详见[使用自己的 i18n 库](/zh/guide/localization#use-your-own-i18n-library)。

## 界面自定义 {#ui-customization}

| 属性            | 类型                  | 默认值 | 说明                                                                                                                                                                                                                                                       |
| --------------- | --------------------- | ------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `customization` | `ViewerCustomization` | -      | 隐藏、锁定或重映射界面外壳的任意部分：功能区选项卡和按钮、文件 > 选项的页面、分区和设置（锁定值、设置默认值）、文件选项卡页面和卡片、右键菜单项、编辑器快捷键、面板、功能（AI、协作）和对话框。与 `hiddenActions` 取并集。传入新对象会替换整个自定义配置。 |

```tsx
<PowerPointViewer
	content={bytes}
	canEdit
	customization={{ ribbon: { hiddenTabs: ['draw'] }, disabledFeatures: ['ai'] }}
/>
```

同样的辅助方法（`hideRibbonTab`、`lockSetting`、`remapShortcut` 等）可在组件 `ref` 上调用，以便在运行时修改，详见 [API 页面](/zh/react/handle#ui-customization)。所有 id、规则和用法见[界面自定义指南](/zh/guide/customization)。

## 协作 {#collaboration}

以下属性启用和控制实时共同编辑。完整流程和 `CollaborationConfig` 结构见[协作](/zh/react/collaboration)。

| 属性                    | 类型                                                         | 默认值 | 说明                                                                                                                                                                                                                                  |
| ----------------------- | ------------------------------------------------------------ | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `collaboration`         | `CollaborationInput`                                         | -      | 提供后启用远程光标、在线状态和 Yjs CRDT 同步。接受旧版 `CollaborationConfig`，或共享的 `CollaborationOptions`（session、user、intent、provider、webrtc）。需要 `yjs` 同级依赖，使用组件自有传输时还需要 `y-websocket` 或 `y-webrtc`。 |
| `onStartCollaboration`  | `(config: CollaborationConfig) => void`                      | -      | 用户从共享对话框创建会话时调用，宿主应使用返回配置设置 `collaboration`。                                                                                                                                                              |
| `onStopCollaboration`   | `() => void`                                                 | -      | 用户从共享对话框停止会话时调用，宿主应清空 `collaboration`。                                                                                                                                                                          |
| `onCollaborationChange` | `(state: CollaborationState) => void`                        | -      | 会话启动或停止，以及其状态、同步标记或参与者变化时，报告共享的 `CollaborationState`。                                                                                                                                                 |
| `shareDefaults`         | `{ roomId?: string; userName?: string; serverUrl?: string }` | -      | 共享对话框字段的默认值，省略时为空。                                                                                                                                                                                                  |

::: info 说明
`collaboration` 是受控属性，组件不会自行启动会话。请在 `onStartCollaboration` 中设置，在 `onStopCollaboration` 中清空。
:::

## 事件、hooks 与宿主扩展 {#events-hooks-and-host-extensions}

| 属性                     | 类型                                 | 说明                                                                                                                    |
| ------------------------ | ------------------------------------ | ----------------------------------------------------------------------------------------------------------------------- |
| `onLoad`                 | `(detail: ViewerLoadDetail) => void` | 每次演示文稿无错误加载完成时触发，附带幻灯片数量和画布尺寸。                                                            |
| `onError`                | `(message: string) => void`          | 每种不同的加载失败触发一次，附带可读消息。                                                                              |
| `onHyperlinkClick`       | `HyperlinkClickHandler`              | 跟随链接之前触发；返回 `false` 可取消默认操作。                                                                         |
| `ribbonAddIns`           | `readonly RibbonAddInTab[]`          | 宿主在帮助之后添加的功能区选项卡。每个命令运行其 `run` 回调，并分派冒泡的 `office-ribbon-add-in` 事件。仅限桌面功能区。 |
| `inspectorInitiallyOpen` | `boolean`                            | 属性面板的初始可见性；省略时保持桌面和移动端默认值。                                                                    |

## 接口示意 {#interface-sketch}

::: info 节选
下面只展示常用成员，并非完整类型。`PowerPointViewerProps` 还包含 `ai`、`accountAuth`、`hiddenActions`、`fonts`、`availableThemes`、`defaultThemeKey`、`onThemeChange`、`onModeChange`、`onSelectionChange`、`onSlideCountChange` 和 `onZoomChange`；上方表格涵盖 `fileName`、`onLoad`、`onError`、`onHyperlinkClick`、`ribbonAddIns`、`inspectorInitiallyOpen`、`defaultLocale`、`availableLocales` 和 `onLocaleChange`。完整列表以包内 `.d.ts` 或 `packages/react/src/viewer/types-ui.ts` 为准。
:::

```ts
interface PowerPointViewerProps {
	content: Uint8Array;
	filePath?: string;
	fileName?: string;

	onDirtyChange?: (isDirty: boolean) => void;
	onContentChange?: (content: Uint8Array) => void;
	onActiveSlideChange?: (slideIndex: number) => void;
	onOpenFile?: () => void;
	onLoad?: (detail: ViewerLoadDetail) => void;
	onError?: (message: string) => void;
	onHyperlinkClick?: HyperlinkClickHandler;
	ribbonAddIns?: readonly RibbonAddInTab[];
	inspectorInitiallyOpen?: boolean;

	canEdit?: boolean;
	initialSlide?: number;
	googleFonts?: false | { baseUrl: string };
	allowExternalContent?: boolean;
	limits?: PptxLoadLimits;
	showToolbar?: boolean;
	showThumbnails?: boolean;
	showCompatibilityToasts?: boolean;
	className?: string;
	authorName?: string;
	smartArt3D?: boolean;
	surfaceChart3D?: boolean;
	barChart3D?: boolean;
	lineChart3D?: boolean;
	areaChart3D?: boolean;
	pieChart3D?: boolean;
	ai?: PptxAiConfig;

	theme?: ViewerTheme;

	collaboration?: CollaborationInput;
	onStartCollaboration?: (config: CollaborationConfig) => void;
	onStopCollaboration?: () => void;
	onCollaborationChange?: (state: CollaborationState) => void;
	shareDefaults?: {
		roomId?: string;
		userName?: string;
		serverUrl?: string;
	};
}
```

## 功能触发方式 {#notes-on-triggering-features}

导航、模式切换、缩放、撤销重做和选择可以通过命令式[句柄](/zh/react/handle)控制，例如 `ref.current.goNext()` 或 `ref.current.setMode('present')`。导出和打印通过内置工具栏和对话框发起，详见[导出](/zh/react/export)。

## 自动保存与恢复 {#autosave-recovery}

`canEdit` 为 true 且提供 `filePath` 时，组件会在有未保存修改时将恢复快照保存到 **IndexedDB**，默认在首次未保存修改两分钟后写入一次。即使关闭页面或发生崩溃，数据仍保留在浏览器中。

### 工作方式 {#how-it-works}

1. 首次未保存修改会启动计时器，默认两分钟，间隔可配置。之后的修改不会延后写入；文档未变化时不写入任何内容。
2. 存在修改时，将当前状态序列化为 `Uint8Array`，存入 IndexedDB 数据库 `pptx-viewer-autosave`，以 `filePath` 为键。
3. 下次加载相同 `filePath`，且幻灯片加载完成后，检查 24 小时内的近期快照，并打开版本历史面板。

### 宿主的职责 {#host-responsibility}

组件**不会**自行跨页面重载持久保存 `filePath`，宿主必须：

1. **记住文件标识**，例如保存在 `localStorage`、URL 参数或后端会话中。
2. 页面重载、组件重新挂载时，**再次传入 `filePath`**。

否则，恢复检查没有可查找的键，无法找到之前的快照。

### 恢复工具函数（共享包） {#recovery-helpers-shared-package}

`pptx-react-viewer` 不重新导出自动保存存储。相关工具函数来自 `ooxml-ui/pptx`（该包的依赖项），可用于构建自定义恢复流程：

```ts
import {
	getAutosaveSnapshot,
	listAutosaveSnapshots,
	deleteAutosaveSnapshot,
	saveAutosaveSnapshot,
} from 'ooxml-ui/pptx';

// List all stored snapshots (without the heavy data blob)
const snapshots = await listAutosaveSnapshots();
// => [{ key: 'report.pptx', timestamp: 1720300000000, size: 524288 }, ...]

// Retrieve a specific snapshot by its key (filePath)
const snapshot = await getAutosaveSnapshot('report.pptx');
if (snapshot) {
	// snapshot.data is a Uint8Array you can pass as `content`
	setContent(snapshot.data);
}

// Delete a snapshot (e.g. after the user dismisses recovery)
await deleteAutosaveSnapshot('report.pptx');
```

### 刷新后重新打开文稿 {#session-restore}

恢复快照用于找回标签页崩溃时的编辑，而更常见的普通**刷新**有专门的工具函数。每个组件都重新导出 `rememberSessionDeck` 和 `restoreSessionDeck`，也可从 `ooxml-ui/pptx` 获取。它们按**浏览器标签页**记住当前打开的文稿，并在下次加载时返回：

```ts
import { rememberSessionDeck, restoreSessionDeck } from 'pptx-react-viewer';

// After the host loads a deck:
await rememberSessionDeck(file.name, bytes);

// On mount, before falling back to your file picker:
const deck = await restoreSessionDeck();
if (deck) {
	setContent(deck.data);
	setFilePath(deck.fileName);
}
```

- **范围限定为当前标签页。** 记录使用 `sessionStorage` 中的 ID 作为键，因此刷新会重新打开文稿，新标签页仍从空白开始，两个标签页中的不同文稿不会相互替换。
- **保留编辑。** `restoreSessionDeck()` 优先选择相同 `fileName` 下更新的自动保存快照，因此编辑过程中刷新会恢复修改后的文稿，而非最初打开的字节。
- **覆盖组件自身的“文件 > 打开”。** 文件后台或最近使用列表会在组件内部切换文稿而不通知宿主，因此这些路径会自行记录会话文稿。
- `forgetSessionDeck()` 删除记录。所有调用均尽力执行，IndexedDB 被阻止或 `sessionStorage` 被分区时，会回退为没有可恢复内容，不抛出错误。

五个演示应用均使用这套流程，因此刷新后仍可保留当前演示文稿。

### 示例：页面重载时恢复 {#example-recovery-on-page-reload}

```tsx
import { getAutosaveSnapshot } from 'ooxml-ui/pptx';
import { PowerPointViewer } from 'pptx-react-viewer';
import { useEffect, useState } from 'react';

const STORAGE_KEY = 'my-app-last-file';

function App() {
	const [content, setContent] = useState<Uint8Array | null>(null);
	const [filePath, setFilePath] = useState('');

	// On mount, check for a recovery snapshot
	useEffect(() => {
		const lastFile = localStorage.getItem(STORAGE_KEY);
		if (!lastFile) return;

		getAutosaveSnapshot(lastFile).then((snapshot) => {
			if (snapshot && Date.now() - snapshot.timestamp < 24 * 60 * 60 * 1000) {
				// Offer recovery (or auto-restore)
				setContent(snapshot.data);
				setFilePath(snapshot.key);
			}
		});
	}, []);

	// When user opens a file, persist the name
	function handleOpen(file: File) {
		localStorage.setItem(STORAGE_KEY, file.name);
		setFilePath(file.name);
		file.arrayBuffer().then((buf) => setContent(new Uint8Array(buf)));
	}

	if (!content) return <FileDropzone onFile={handleOpen} />;

	return <PowerPointViewer content={content} filePath={filePath} canEdit />;
}
```

### autosave 属性和自动保存开关，谁优先 {#autosave-policy}

宿主属性和用户开关都能控制自动保存，五个组件使用共享函数 `resolveAutosaveActivation` 统一裁决：

> **`autosave` 属性决定策略上限，标题栏开关表达该范围内的用户偏好。**

| autosave 属性 | 实际行为                                       | 开关                         |
| ------------- | ---------------------------------------------- | ---------------------------- |
| 省略          | 允许自动保存，由用户开关决定，默认**开启**。   | 可用。                       |
| `true`        | 与省略相同，宿主允许，由用户决定。             | 可用。                       |
| `false`       | 关闭自动保存，不写快照，也不在加载时提示恢复。 | **不可操作**，状态不能切换。 |

显式属性表示应用允许的范围，用户开关不能越过这个范围，因此 `autosave={false}` 也会禁用开关。还必须满足两个前提：`canEdit` 为 true，并且设置了 `filePath`，否则没有可写内容或保存位置。

保存间隔遵循同样规则：显式 `autosaveIntervalMs` 是宿主策略，按传入值执行；省略时，使用**文件 > 选项 > 保存 > 每 N 分钟保存自动恢复信息**中的用户设置，默认两分钟。

自动保存默认开启，以便默认提供崩溃恢复。需要退出此行为时传入 `autosave={false}`。

### 恢复快照 {#recovering-a-snapshot}

文稿加载完成后，如果相同 `filePath` 存在 24 小时内的快照，会显示 **是否恢复未保存的更改？** 对话框。恢复会加载快照字节，丢弃会删除快照。如果当前标签页已接收该快照，例如宿主通过 `restoreSessionDeck` 恢复过，则不会再次提示。

### 前提条件与状态反馈 {#requirements-and-status-feedback}

条件不满足时，自动保存会停用，并在标题栏显示原因：

| 条件               | 显示状态                       | 原因                                           |
| ------------------ | ------------------------------ | ---------------------------------------------- |
| 未提供 `filePath`  | 自动保存已禁用：未提供文件路径 | 需要稳定的键保存快照，传入 `filePath` 可启用。 |
| 用户关闭自动保存   | 自动保存已关闭                 | 用户通过标题栏开关主动停用。                   |
| `autosave={false}` | 此应用已关闭自动保存           | 宿主禁止此功能，开关不可操作。                 |
| `canEdit` 为 false | 自动保存已禁用：演示文稿为只读 | 自动保存仅在编辑模式下适用。                   |

补齐条件，例如设置 `filePath` 后，状态会自动回到 `'idle'` 并启动计时器。

`AutosaveStatus` 类型如下：

```ts
type AutosaveStatus =
	| { state: 'idle' }
	| { state: 'disabled'; reason: string }
	| { state: 'saving' }
	| { state: 'saved'; timestamp: number }
	| { state: 'error'; message: string };
```

`reason` 为 `'autosave_host_off'`、`'autosave_toggle_off'`、`'no_file_path'` 或 `'read_only'`，共享的 `autosaveDisabledReasonKey` 分别将其映射到 `pptx.autosave.disabledByHost`、`pptx.autosave.disabledToggleOff`、`pptx.autosave.disabledNoFilePath` 和 `pptx.autosave.disabledReadOnly`。
