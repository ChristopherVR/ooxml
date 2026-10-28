---
title: Hooks
description: 了解 PowerPointViewer 的 hooks 架构、/viewer 提供的稳定公共 hooks，以及 /internals 暴露的完整内部实现。
---

# Hooks {#hooks}

`PowerPointViewer` 是精简的 `forwardRef` 协调组件，绝大部分逻辑由 `PowerPointViewer.tsx` 中组合的 **80 多个自定义 hook** 实现，视觉组件主要负责展示。状态完全保存在 React hooks 中，没有外部状态库。

::: info 公共与内部接口
大部分 hook 属于**内部架构**，依赖特定组合顺序和共享输入。经过筛选的小部分从 `pptx-react-viewer/viewer` 导出，遵循正常的语义化版本兼容保证。**完整集合**可从 `pptx-react-viewer/internals` 导入，但内部构建块不受该保证约束，应优先使用稳定导出。详见[完整 Hooks 参考](/zh/react/hooks-reference)。
:::

## 内部架构 {#architecture-internal}

下表说明组件内部如何连接。虽然这些 hook 可以导入，但它们依赖特定组合顺序和输入，应将此表视为架构说明，而非 API 契约。

| Hook                     | 职责                                        |
| ------------------------ | ------------------------------------------- |
| `useViewerState`         | 组合核心状态和 UI 状态。                    |
| `useViewerCoreState`     | 幻灯片、选择、画布尺寸和模式等文档状态。    |
| `useViewerUIState`       | 面板可见性、对话框和工具栏标记。            |
| `useDerivedSlideState`   | 计算可见索引、节和母版伪幻灯片。            |
| `useEditorHistory`       | 撤销重做快照栈，在指针交互期间延迟捕获。    |
| `useZoomViewport`        | 缩放、适应宽度和视口 DOM ref。              |
| `useEditorOperations`    | 将全部编辑操作组合为一个结果。              |
| `useLoadContent`         | 挂载时通过 `PptxHandler` 解析 PPTX 缓冲区。 |
| `useContentLifecycle`    | 内容同步、未保存修改跟踪和恢复检测。        |
| `usePresentationMode`    | 放映导航、动画和切换。                      |
| `useExportHandlers`      | PNG、SVG、PDF、GIF、视频和 PPTX 导出。      |
| `usePrintHandlers`       | 打印对话框和布局。                          |
| `useInsertElements`      | 插入形状、图片、文本框、表格和图表。        |
| `useElementManipulation` | 移动、缩放、旋转和删除元素。                |
| `useSlideManagement`     | 添加、删除、复制、重排和隐藏幻灯片。        |
| `useTableOperations`     | 插入和删除行列，合并和拆分单元格。          |
| `usePointerHandlers`     | 处理画布的鼠标和触控事件。                  |
| `useKeyboardShortcuts`   | 定义快捷键。                                |
| `useViewerIntegration`   | 顶层整合 I/O、导出、打印、指针和生命周期。  |

此外还有剪贴板、批注、节、自动保存、字体注入、恢复、主题和放映子 hook 等数十项。按职责分组的完整列表见[完整 Hooks 参考](/zh/react/hooks-reference)。

内部 hooks 通常需要连接大量状态。例如 `useEditorHistory` 接收完整编辑状态和 8 个状态 setter，返回 `{ canUndo, canRedo, undoLabel, redoLabel, handleUndo, handleRedo, resetHistory, markDirty, buildHistorySnapshot }`。这种结构适合组件内部组合；独立实现撤销重做时，可以直接修改 `PptxData` 并保存快照，详见[编程编辑](/zh/core/editing)。

## 公共 hooks {#public-hooks}

以下 hooks 从 `pptx-react-viewer/viewer` 导出，可按需使用并支持 tree-shaking。`useViewerBuildingBlocks` 也从包根入口导出，同时导出的还有独立的 `Toolbar`、`SlideCanvas` 和 `SlidesPane` 组件。

```tsx
import { useThemeSwitching, useCollaborativeState } from 'pptx-react-viewer/viewer';
```

### 自定义编辑器的协作 {#custom-editor-collaboration}

`useViewerBuildingBlocks` 接受与 `PowerPointViewer` 相同的可选 `collaboration` 配置，包括[宿主自有的 Yjs 会话](/zh/guide/host-owned-collaboration)。它返回可用于自定义状态栏的 `collaboration` 连接和在线状态，并在 `canvasProps` 中包含光标和远程选择叠加层，无需额外的协作 provider。

```tsx
import { SlideCanvas, Toolbar, useViewerBuildingBlocks } from 'pptx-react-viewer';

function Editor({ content, collaboration }) {
	const blocks = useViewerBuildingBlocks({ content, collaboration, canEdit: true });
	return (
		<>
			<Toolbar {...blocks.toolbarProps} />
			<SlideCanvas {...blocks.canvasProps} />
		</>
	);
}
```

该 hook 使用内置编辑器的文档同步、就绪检查和写回路径。`role: 'viewer'` 会禁用编辑。这不会改变其现有的撤销语义，也不会增加单独的宿主批注集成。

仅在配置了协作时才会增加源加载和错误检查。未配置协作时，该 hook 保持原有的宿主授权行为，包括用 `content: null` 表示空白文档。活动的协作会话同样遵循宿主的编辑权限和就绪状态；仅有已连接的传输层并不授予编辑权限。宿主需要为返回的状态类型命名时，可从包根入口使用 `CollaborationContextValue`。

### 幻灯片缩略图栏 {#slide-thumbnail-rail}

`slidesPaneProps` 驱动 `SlidesPane`，即 `PowerPointViewer` 在画布旁显示的幻灯片缩略图栏。它列出所有幻灯片（演示文稿包含节时按节分组）并突出显示当前幻灯片；点击缩略图会将画布切换到该幻灯片，不会把演示文稿标记为已修改。启用 `canEdit` 后，它还可以添加、复制、删除、隐藏、重新排序和分组幻灯片，使用与完整编辑器相同的操作、撤销历史和协作同步。

```tsx
import { SlideCanvas, SlidesPane, Toolbar, useViewerBuildingBlocks } from 'pptx-react-viewer';

function Editor({ content }) {
	const blocks = useViewerBuildingBlocks({ content, canEdit: true });
	return (
		<>
			<Toolbar {...blocks.toolbarProps} />
			<div style={{ display: 'flex', height: 600 }}>
				<SlidesPane {...blocks.slidesPaneProps} />
				<div style={{ flex: 1, minWidth: 0, position: 'relative' }}>
					<SlideCanvas {...blocks.canvasProps} />
				</div>
			</div>
		</>
	);
}
```

缩略图栏的展开状态就是工具栏的侧边栏开关：折叠任一处都会隐藏缩略图栏。幻灯片右键菜单中的“版式”项会将该幻灯片设为当前幻灯片，并调用可选的 `onOpenLayoutGallery({ slideIndex, x, y })`，因为这些构建块不渲染版式库。`SlidesPane` 需要完整的 `styles.css`，而不是 `canvas.css`。

### 打开指定幻灯片 {#opening-on-a-slide}

`useViewerBuildingBlocks` 接受与 `PowerPointViewer` 相同的 `initialSlide` 选项：每次加载后显示的幻灯片，索引从 0 开始，自动限制在有效范围内。它在首次显示已加载演示文稿的那次渲染中生效，因此自定义界面不会先显示第 1 张幻灯片，也无需在加载后调用 `goTo()` 或在跳转完成前隐藏画布。

```tsx
const blocks = useViewerBuildingBlocks({ content, initialSlide: 4 });
```

之后修改 `initialSlide` 只对下一次加载生效，不会移动已打开的演示文稿；要切换当前幻灯片，请使用句柄的 `goTo(index)`。

### `useThemeSwitching` {#usethemeswitching}

切换已加载**文档本身**的 PowerPoint 主题，即 OOXML 颜色和字体方案，不是组件界面主题，两者区别见[主题配置](/zh/guide/theming)。它使用手动加载得到的同一组 `PptxHandler` 和 `PptxData`。

```ts
interface UseThemeSwitchingInput {
	handlerRef: RefObject<PptxHandler | null>;
	data: PptxData | null;
	onDataChange: (newData: PptxData) => void;
	onThemeChanged?: (preset: PptxThemePreset) => void;
}

interface ThemeSwitchingResult {
	presets: readonly PptxThemePreset[]; // built-in presets (office, facet, ion, ...)
	switchToPreset: (preset: PptxThemePreset) => Promise<void>;
	switchToCustom: (
		colorScheme: PptxThemeColorScheme,
		fontScheme?: PptxThemeFontScheme,
		themeName?: string,
	) => Promise<void>;
	currentPreset: PptxThemePreset | undefined; // preset matching the current theme, if any
}
```

`switchToPreset` 同时更新内存中的 ZIP 和解析数据中已求值的元素颜色，使修改在 `save()` 后仍保留。

```tsx
function ThemePicker({ handlerRef, data, setData }: Props) {
	const { presets, switchToPreset, currentPreset } = useThemeSwitching({
		handlerRef,
		data,
		onDataChange: setData,
	});

	return (
		<div>
			{presets.map((preset) => (
				<button
					key={preset.id}
					onClick={() => switchToPreset(preset)}
					aria-pressed={preset.id === currentPreset?.id}
				>
					{preset.name}
				</button>
			))}
		</div>
	);
}
```

### 协作 hooks {#collaboration-hooks}

用于自定义协作界面或自行控制同步，详见[协作](/zh/react/collaboration)。需要可选依赖 `yjs` 和 `y-websocket`，它们动态加载，未使用时可被 tree-shaking 移除。各 hook 使用与组件 `collaboration` 属性相同的 `CollaborationConfig`，包括 `roomId`、`serverUrl`、`userName`，以及可选的 `transport`（`'websocket' | 'webrtc'`）、`signaling`、`userColor`、`userAvatar`、`authToken`、`role`、`sessionIntent`，和写入者使用的 `onWriteBack`、`writeBackDebounceMs`。

| Hook                      | 签名（输入 => 结果）                                                                                                                       |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| `useYjsProvider`          | `{ config?: CollaborationConfig }` => `{ status, awareness, doc, clientId, synced, retry }`                                                |
| `usePresenceTracking`     | `{ awareness, localClientId, userName, userColor, userAvatar?, role?, canvasWidth, canvasHeight }` => `{ remoteUsers, broadcastPresence }` |
| `useCollaborativeState`   | `{ config?, canvasWidth, canvasHeight }` => `CollaborationContextValue \| null`                                                            |
| `useCollaborativeHistory` | `{ localClientId, handleUndo, handleRedo, canUndo, canRedo }` => 包装后的相同四项，仅对本地修改执行撤销                                    |

`useCollaborativeState` 是内置 `CollaborationProvider` 的组合入口，管理传输层 `useYjsProvider` 和在线状态 `usePresenceTracking`。`config` 为 `undefined` 时返回 `null`，hooks 保持休眠，组件树结构保持稳定。

```tsx
import { useCollaborativeState } from 'pptx-react-viewer/viewer';

function PresenceBar({ roomId, userName }: { roomId: string; userName: string }) {
	const collab = useCollaborativeState({
		config: { roomId, serverUrl: 'wss://collab.example.com', userName },
		canvasWidth: 960,
		canvasHeight: 540,
	});

	if (!collab) return null;
	return (
		<span>{collab.status === 'connected' ? `${collab.connectedCount} online` : collab.status}</span>
	);
}
```

`useYjsProvider` 是轻量传输层，延迟导入 Yjs 包，创建 `Y.Doc` 和 Provider。连接失败超时后进入 `status: 'error'`，可通过 `retry()` 恢复。它还提供 `synced`，让后加入的用户等房间文档到达后再允许本地写入。

同时导出的还有 `CollaborationProvider`、`RemoteUserCursors`、`UserAvatarBar` 和 `CollaborationStatusIndicator`。

### 观众窗口工具 {#audience-window-helpers}

以下工具不是 hooks，但从 `pptx-react-viewer/viewer` 导出，用于演讲者和观众窗口之间的流程：`isAudienceTab`、`loadAudienceContent`、`storeAudienceContent`、`clearAudienceContent` 和 `parseAudienceNonce`。

## 直接使用内部 hook {#using-an-internal-hook-directly}

公共 hooks 无法满足需求时，可以从 `pptx-react-viewer/internals` 导入全部内部 hook：

```tsx
import { useEditorHistory, useViewerState } from 'pptx-react-viewer/internals';
```

::: warning 内部构建块
`pptx-react-viewer/internals` 原样导出组件内部使用的 hooks，**不受语义化版本兼容保证约束**。签名、行为、名称乃至是否存在，都可能在不升级主版本的情况下改变。请优先使用属性、句柄或 `/viewer` 的公共 hooks，仅在高级接入需要时使用内部入口，并锁定精确版本。
:::

完整列表见[完整 Hooks 参考](/zh/react/hooks-reference)，整体结构见[组件概览](/zh/react/#hooks-based-architecture)。
