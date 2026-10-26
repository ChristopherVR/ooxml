---
title: 界面自定义
description: 通过代码隐藏、锁定、重映射或替换查看器界面的任何部分，包括功能区标签页和按钮、“文件 > 选项”页面、分区和设置、“文件”选项卡、右键菜单、键盘快捷键、面板、功能、对话框和能力，或将设置对话框、检查器、面板和“文件”页面换成你自己的界面。使用一个与框架无关的对象，在 React、Vue、Angular、Svelte 和 Vanilla JS 中行为一致。
---

# 界面自定义

将查看器的 `inspectorInitiallyOpen` 选项设为 `false`，即可在启动时关闭属性检查器。面板开关和功能区组启动器仍可打开它。省略该选项则保留绑定的桌面/移动端默认行为；之后对这个初始选项的修改不会重置用户的面板状态。

使用共享的自定义对象隐藏无选中内容时检查器中的单个卡片：

```ts
const customization: ViewerCustomization = {
	inspector: { hiddenCards: ['presentation', 'theme', 'themeOverride', 'slideSize'] },
};
```

卡片 id 为 `presentation`、`theme`、`themeOverride`、`slideSize`、`slideTransition`、`notesHandout`、`document`、`tags`、`slide`、`slideBackground` 和 `templateBackground`。每个 id 会在绑定提供该卡片的所有位置将其隐藏。`updateCustomization({ inspector: { hiddenCards: [] } })` 可实时恢复这些卡片。卡片可见性按每个查看器独立设置，与检查器的打开和关闭无关。

每个绑定都渲染相同的 PowerPoint 风格界面：功能区、“文件”选项卡、“文件 > 选项”（设置）对话框、右键菜单、键盘映射、幻灯片栏、检查器、状态栏以及若干对话框。大多数嵌入场景并不需要全部这些内容。展台不需要任何编辑界面；企业门户希望去掉 AI 助手，并将作者名固定为已登录用户；教学工具则只需要三个标签页的功能区。

这一切都由一个对象 `ViewerCustomization` 完成，你把它传给查看器，并可在运行时修改：

```ts
import type { ViewerCustomization } from 'pptx-react-viewer'; // 或 vue / angular / svelte / vanilla

const customization: ViewerCustomization = {
	ribbon: {
		hiddenTabs: ['draw', 'record'],
		hiddenGroups: ['insert.media'],
		hiddenButtons: ['broadcast', 'home.font.changeCase'],
	},
	options: {
		hiddenPages: ['trust', 'addIns'],
		locked: { 'general.userName': 'Ada Lovelace' },
		defaults: { 'advanced.showGrid': true },
	},
	contextMenu: { hiddenElementCommands: ['save-as-picture'] },
	keyboard: { disabled: ['newSlide'], remap: { duplicate: 'Mod+Shift+D' } },
	hiddenPanels: ['notes'],
	disabledFeatures: ['ai'],
	hiddenDialogs: ['broadcast'],
	hiddenExportFormats: ['video', 'gif'],
	capabilities: { print: false },
	regions: { 'dialog.options': 'acme-settings' }, // 你自己的设置视图
};
```

该对象、其类型、id 目录和辅助方法在全部五个绑定中完全一致。逻辑只在 `ooxml-ui/pptx`（`src/ui/src/pptx/render/customization/`）中存在一份，因此在 React 中可用的自定义在 Svelte 中的行为完全相同。

[[toc]]

## 可配置的内容 {#what-you-can-configure}

下面的一切要么是 `customization` 的字段（本指南），要么是查看器组件的属性（绑定自己的指南）。自定义字段如下：

| 你想要                                                                     | 字段                                                                                                            |
| -------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| 隐藏功能区标签页、组、控件和工具栏按钮                                     | `ribbon.hiddenTabs`, `ribbon.hiddenGroups`, `ribbon.hiddenButtons`                                              |
| 添加你自己的功能区标签页                                                   | `ribbonAddIns` 属性（[见下文](#host-ribbon-tabs-ribbonaddins)）                                                 |
| 隐藏、锁定“文件 > 选项”的页面、分区和设置，或为其设置默认值                | `options.hiddenPages`, `options.hiddenSections`, `options.hiddenSettings`, `options.locked`, `options.defaults` |
| 完全移除“文件 > 选项”                                                      | `hiddenDialogs: ['options']`                                                                                    |
| 用你自己的设置视图替换“文件 > 选项”                                        | `regions['dialog.options']`                                                                                     |
| 隐藏“文件”选项卡的页面和卡片                                               | `backstage.hiddenPages`, `backstage.hiddenCards`                                                                |
| 添加你自己的“文件”选项卡页面                                               | `backstage.extraPages`                                                                                          |
| 替换“文件”选项卡的某个页面                                                 | `regions['backstage.<page>']`                                                                                   |
| 隐藏或添加右键菜单命令                                                     | `contextMenu.*`                                                                                                 |
| 禁用或重映射快捷键                                                         | `keyboard.*`                                                                                                    |
| 隐藏面板（检查器、幻灯片栏、备注、状态栏、标题栏、选择窗格、查找和替换等） | `hiddenPanels`                                                                                                  |
| 替换检查器、幻灯片栏、备注或状态栏                                         | `regions.inspector`, `regions.slidesPane`, `regions.notes`, `regions.statusBar`                                 |
| 隐藏检查器卡片                                                             | `inspector.hiddenCards`                                                                                         |
| 关闭 AI、协作、批注、放映模式、编辑顶点                                    | `disabledFeatures`                                                                                              |
| 移除对话框及所有打开它们的入口                                             | `hiddenDialogs`                                                                                                 |
| 替换共享、广播或打印对话框的主体                                           | `regions['dialog.share']`, `regions['dialog.broadcast']`, `regions['dialog.print']`                             |
| 移除导出格式                                                               | `hiddenExportFormats`                                                                                           |
| 移除绘图工具                                                               | `hiddenDrawingTools`                                                                                            |
| 将演示文稿设为只读，或禁止导出、打印、共享、复制、放映或批注               | `capabilities`                                                                                                  |

## 概念 {#concepts}

### 隐藏、锁定、默认值 {#hide-lock-default}

对“文件 > 选项”中的某个设置可以做三种不同的事，它们可以组合使用：

| 你想要                         | 使用                                                          |
| ------------------------------ | ------------------------------------------------------------- |
| 用户永远看不到该设置           | `options.hiddenSettings: ['general.userInitials']`            |
| 该设置有固定值，用户无法更改   | `options.locked: { 'general.userName': 'Ada' }`（以只读渲染） |
| 两者兼有：固定值且完全没有控件 | 将该 id 同时放入 `locked` **和** `hiddenSettings`             |
| 不同的起始值，但用户仍可更改   | `options.defaults: { 'advanced.showGrid': true }`             |

- **锁定**的值会立即强制写入选项存储，此后无论来自哪条路径（对话框、写入同一选项的功能区开关、重置），对它的写入都会被忽略。控件以禁用状态渲染，并带有工具提示“此设置由您的组织管理”。锁定的值绝不会写入用户保存的偏好，因此移除锁定后，用户会恢复自己的值。
- **默认值**会替换尚未保存自己选择的用户的内置默认值，并成为“重置”所返回的值。宿主默认值同样不会被持久化：在后续版本中更改默认值，所有从未改动该设置的用户都会跟随，而保留了自己选择的用户则不受影响。

### 派生规则 {#derived-rules}

有些 id 会同时关闭多个入口，因此你不必列出每个打开某个对话框的按钮。这些规则在共享的解析器中统一应用一次：

- `hiddenDialogs: ['share']` 会移除共享按钮、“文件 > 共享”页面和共享卡片。`['print']` 移除“文件 > 打印”和打印卡片；`['export']` 移除导出按钮和“文件 > 导出”；`['options']` 移除“文件 > 选项”。
- `disabledFeatures: ['collaboration']` 隐藏共享和广播对话框（因此上述针对它们的内容也一并生效）。
- `disabledFeatures: ['ai']` 移除 AI 开关和面板、“询问 AI” / “用 AI 修复”右键菜单项以及“文件 > 选项”的 AI 页面，即使宿主传入了 `ai` 配置也是如此。
- `disabledFeatures: ['comments']` 移除“添加批注”右键菜单项以及“插入”和“审阅”选项卡的批注组。
- `disabledFeatures: ['presentMode']` 移除“幻灯片放映”功能区标签页以及 F5 / Shift+F5 开始放映快捷键。
- `hiddenExportFormats` 移除“文件 > 导出”中的单个卡片；六个全部隐藏时，导出按钮和页面也会被移除。
- 隐藏面板会移除打开它的功能区控件和快捷键：`findReplace` 会连带移除“开始 > 查找 / 替换”和 Ctrl+F / Ctrl+H，`selectionPane` 移除“视图 > 选择窗格”，`accessibilityPane` 移除“审阅 > 检查辅助功能”和“帮助 > 辅助功能检查器”，`comparePane` 移除“审阅 > 比较”，`notes` 移除备注按钮。
- 隐藏对话框对其入口同样如此：`hyperlink` 移除“插入 > 链接”、链接菜单项和 Ctrl+K；`equation` 移除“插入 > 公式”；`pasteSpecial` 移除其快捷键；`slideTemplates` 移除“开始 > 新建幻灯片”中的幻灯片模板项；`passwordProtection` 和 `signatures` 移除其“文件 > 信息”卡片。
- `capabilities` 会展开为上面的列表（参见[能力](#capabilities)）。
- 失去所有控件的选项分区会被移除；失去所有分区的选项页面也会被移除。当某组的第一项被隐藏时，菜单会自行修复分隔线。

### 属性与命令式 API {#props-and-the-imperative-api}

将该对象作为 `customization` 属性（输入、选项）传入以预先设置，并使用组件句柄上的命令式辅助方法在查看器运行时修改它。每次更改都会立即重新渲染受影响的界面，无需重新挂载。

- **属性是替换**：当你给查看器一个新的 `customization` 对象时，它就成为完整的自定义，之前通过辅助方法所做的修改都会被丢弃。除非你有意替换，否则请保持该对象稳定（用 memo 缓存，或保存在状态中）。
- **辅助方法是合并**：`hideRibbonTab('draw')` 会向现有内容中添加一个 id。`updateCustomization(patch)` 合并一层深度（每个分区逐字段合并；你传入的列表会替换该列表）。

### 向后兼容：`hiddenActions` 和自定义功能区 {#back-compat-hiddenactions-and-customize-ribbon}

较早的 `hiddenActions` 属性仍然有效。它会与 `ribbon.hiddenTabs` 和 `ribbon.hiddenButtons` 取并集，因此你可以按自己的节奏迁移。用户自己在“文件 > 选项 > 自定义功能区”中的选择会在其上再取并集：只有宿主和用户都没有隐藏的标签页才会显示。

### 功能区组和控件 {#ribbon-groups-and-controls}

在标签页层级之下，每个功能区组和组内的每个控件都在一个共享目录中拥有稳定的 id：组为 `<tab>.<group>`，控件为 `<tab>.<group>.<control>`，按 PowerPoint 自身的标题命名（`home.font`、`home.font.bold`、`insert.media.media`、`view.show.ruler`）。因选择而出现的上下文标签页（`shapeFormat`、`pictureFormat`、`tableDesign`、`chartDesign`、`smartArtDesign`）以同样方式寻址，其标签页 id 可用于 `ribbon.hiddenTabs`。

```ts
const customization: ViewerCustomization = {
	ribbon: {
		hiddenGroups: ['home.editing', 'insert.media'],
		hiddenButtons: ['home.font.changeCase', 'home.paragraph.columns', 'mergeShapes'],
		hiddenTabs: ['chartDesign'],
	},
};
```

`ribbon.hiddenButtons` 同时接受两种 id：较早的顶层工具栏按钮（`share`、`zoom` 等）和目录中的功能区控件。`mergeShapes` / `home.arrange.mergeShapes` 以及 `crop` / `home.arrange.crop` 指向同一个控件，因此两种写法都能将其隐藏。两种裁剪写法也会隐藏第二个裁剪入口 `pictureFormat.size.crop`（图片格式 > 大小）；指定 `pictureFormat.size` 或 `pictureFormat.size.crop` 则只隐藏那一个。完整列表见下方的[参考](#ribbon-groups-ribbon-hiddengroups)。

每个绑定都会为其功能区标记添加 `data-ribbon-group` 和 `data-ribbon-control`，并渲染一份由共享模型根据解析后的自定义生成、限定在该查看器范围内（`data-pptx-ribbon-scope`）的样式表，因此隐藏的组或控件在全部五个绑定中都以相同方式从布局和无障碍树中移除。这些属性同样是你自己的样式或测试的稳定挂钩。

### 宿主功能区标签页（`ribbonAddIns`） {#host-ribbon-tabs-ribbonaddins}

自定义是移除内容；`ribbonAddIns` 则像 Office 加载项一样，在“帮助”之后添加你自己的标签页。它是每个绑定上的独立属性（React、Svelte 和 vanilla 选项中为 `ribbonAddIns`，Vue 中为 `:ribbon-add-ins`，Angular 中为 `[ribbonAddIns]`；vanilla 实例还有 `setRibbonAddIns(tabs)`），并接受与 Word、Excel 和 Visio 编辑器相同的普通描述符：

```tsx
const tabs = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [
			{
				label: 'Export',
				commands: [
					{ id: 'export', label: 'Export', icon: 'save', run: () => exportReport() },
					{ id: 'options', label: 'Options', size: 'small' },
				],
			},
		],
	},
];

<PowerPointViewer content={bytes} ribbonAddIns={tabs} />;
```

选中某个命令后，会运行其 `run` 回调，然后从查看器内部派发一个冒泡的 `office-ribbon-add-in` DOM 事件（`detail: { tab, command }`），因此在查看器容器上放一个监听器即可听到所有宿主命令。它没有对应的各绑定回调属性。

- 标签页不能使用内置标签页或上下文标签页的 id；这样的标签页会被忽略。`label` 是你自己的文本，不会被翻译。
- `size: 'small'` 的命令以三个一列排布。PowerPoint 功能区没有宿主命令的下拉菜单，因此带 `items` 的命令会被绘制为其各项，即以父项标签为标题的小命令。
- 再次传入相等的描述符（框架每次渲染都会这样做）会保留该标签页的控件，只替换回调。移除正在显示的标签页会使功能区返回“开始”。
- 宿主标签页仅显示在桌面功能区中；手机工具栏及其菜单面板不会列出它们。
- `hiddenActions` 和 `ribbon.hiddenTabs` 不适用于宿主标签页：要隐藏某个标签页，只需将其从列表中去掉。

## 命令式 API {#the-imperative-api}

每个绑定都在其组件句柄上公开这些方法（React `ref`、Vue 模板 ref、Angular 组件实例、Svelte `bind:this`、Vanilla 实例）：

| 方法                                                                    | 作用                                                           |
| ----------------------------------------------------------------------- | -------------------------------------------------------------- |
| `getCustomization()`                                                    | 当前的 `ViewerCustomization`（快照；请勿修改它）。             |
| `setCustomization(c)`                                                   | 替换整个自定义。                                               |
| `updateCustomization(patch)`                                            | 合并部分自定义。                                               |
| `resetCustomization()`                                                  | 恢复默认界面。                                                 |
| `hideRibbonTab(id)` / `showRibbonTab(id)`                               | 切换一个功能区标签页（上下文标签页 id 会使该标签页不再出现）。 |
| `hideRibbonGroup(id)` / `showRibbonGroup(id)`                           | 切换标签页内的一个组（`home.font`）。                          |
| `hideToolbarButton(id)` / `showToolbarButton(id)`                       | 切换一个工具栏按钮、控件簇或功能区控件。                       |
| `hideRibbonControl(id)` / `showRibbonControl(id)`                       | 切换一个功能区控件（`home.font.bold`）。                       |
| `hideOptionsPage(id)` / `showOptionsPage(id)`                           | 切换“文件 > 选项”的一个页面。                                  |
| `hideOptionsSection(id)` / `showOptionsSection(id)`                     | 切换选项页面的一个分区。                                       |
| `hideSetting(id)` / `showSetting(id)`                                   | 切换一个设置。                                                 |
| `lockSetting(id, value, hidden?)`                                       | 将设置固定为 `value`；`hidden: true` 同时移除其控件。          |
| `unlockSetting(id)`                                                     | 移除锁定。                                                     |
| `setSettingDefault(id, value)`                                          | 设置宿主默认值（`undefined` 表示清除）。                       |
| `hideBackstagePage(id)` / `showBackstagePage(id)`                       | 切换“文件”选项卡的一个页面。                                   |
| `hideBackstageCard(id)` / `showBackstageCard(id)`                       | 切换“文件”选项卡的一个操作卡片。                               |
| `hideContextMenuCommand(id)` / `showContextMenuCommand(id)`             | 切换一个元素右键菜单项。                                       |
| `hideCanvasContextMenuCommand(id)` / `showCanvasContextMenuCommand(id)` | 切换一个空白画布右键菜单项。                                   |
| `disableShortcut(id)` / `enableShortcut(id)`                            | 切换一个编辑器快捷键。                                         |
| `remapShortcut(id, chords)`                                             | 将命令移到新的组合键；`undefined` 恢复内置组合键。             |
| `setPanelVisible(id, visible)`                                          | 显示或隐藏一个界面区域。                                       |
| `setFeatureEnabled(id, enabled)`                                        | 开启或关闭一个功能领域。                                       |
| `setDialogAvailable(id, available)`                                     | 允许或移除一个对话框及其入口。                                 |
| `setCapability(id, allowed)`                                            | 切换一项能力；`undefined` 恢复默认值。                         |
| `setRegion(id, renderer)`                                               | 用你的界面替换一个区域；`undefined` 恢复内置界面。             |
| `addBackstagePage(page)` / `removeBackstagePage(id)`                    | 添加（或按 id 替换）和移除宿主“文件”选项卡页面。               |

### 快捷键组合 {#shortcut-chords}

组合键的形式为 `Modifier+Modifier+Key`：修饰键为 `Mod`（Windows 和 Linux 上是 Ctrl，macOS 上是 Cmd）、`Ctrl`、`Meta`、`Alt` 和 `Shift`；键为 `KeyboardEvent.key` 值，例如 `D`、`Delete`、`ArrowLeft` 或 `F2`（`+` 写作 `Plus`）。字母不区分大小写。被重映射的命令不再响应其内置组合键。重映射后的组合键保留所有内置保护：在只读查看器中、放映期间、用户在文本框中输入时，或（对于选区命令）没有选中任何内容时，它们都不会触发。`nudge` 和 `escape` 可以禁用，但不能重映射。

## 各绑定的用法 {#per-binding-usage}

::: code-group

```tsx [React]
import { PowerPointViewer } from 'pptx-react-viewer';
import type { PowerPointViewerHandle, ViewerCustomization } from 'pptx-react-viewer';
import { useMemo, useRef } from 'react';

export function Deck({ bytes }: { bytes: Uint8Array }) {
	const viewer = useRef<PowerPointViewerHandle>(null);
	// 使用 memo 缓存：每次渲染都生成新对象会替换辅助方法所做的修改。
	const customization = useMemo<ViewerCustomization>(
		() => ({ ribbon: { hiddenTabs: ['draw'] }, disabledFeatures: ['ai'] }),
		[],
	);
	return (
		<>
			<button onClick={() => viewer.current?.hideRibbonTab('insert')}>Hide Insert</button>
			<PowerPointViewer ref={viewer} content={bytes} canEdit customization={customization} />
		</>
	);
}
```

```vue [Vue]
<script setup lang="ts">
import { ref } from 'vue';
import { PowerPointViewer } from 'pptx-vue-viewer';
import type { PowerPointViewerExpose, ViewerCustomization } from 'pptx-vue-viewer';

defineProps<{ bytes: Uint8Array }>();
const viewer = ref<PowerPointViewerExpose | null>(null);
const customization: ViewerCustomization = {
	ribbon: { hiddenTabs: ['draw'] },
	disabledFeatures: ['ai'],
};
</script>

<template>
	<button @click="viewer?.hideRibbonTab('insert')">Hide Insert</button>
	<PowerPointViewer ref="viewer" :content="bytes" can-edit :customization="customization" />
</template>
```

```ts [Angular]
import { Component, viewChild } from '@angular/core';
import { PowerPointViewerComponent } from 'pptx-angular-viewer';
import type { ViewerCustomization } from 'pptx-angular-viewer';

@Component({
	selector: 'app-deck',
	imports: [PowerPointViewerComponent],
	template: `
		<button (click)="viewer()?.hideRibbonTab('insert')">Hide Insert</button>
		<pptx-viewer #viewer [content]="bytes" [canEdit]="true" [customization]="customization" />
	`,
})
export class DeckComponent {
	bytes: Uint8Array | null = null;
	readonly viewer = viewChild<PowerPointViewerComponent>('viewer');
	readonly customization: ViewerCustomization = {
		ribbon: { hiddenTabs: ['draw'] },
		disabledFeatures: ['ai'],
	};
}
```

```svelte [Svelte]
<script lang="ts">
	import { PowerPointViewer } from 'pptx-svelte-viewer';
	import type { ViewerCustomization } from 'pptx-svelte-viewer';

	let { bytes }: { bytes: Uint8Array } = $props();
	let viewer: ReturnType<typeof PowerPointViewer> | undefined = $state();
	const customization: ViewerCustomization = {
		ribbon: { hiddenTabs: ['draw'] },
		disabledFeatures: ['ai'],
	};
</script>

<button onclick={() => viewer?.hideRibbonTab('insert')}>Hide Insert</button>
<PowerPointViewer bind:this={viewer} source={bytes} editable {customization} />
```

```ts [Vanilla]
import { createPptxViewer } from 'pptx-vanilla-viewer';
import type { ViewerCustomization } from 'pptx-vanilla-viewer';

const customization: ViewerCustomization = {
	ribbon: { hiddenTabs: ['draw'] },
	disabledFeatures: ['ai'],
};
const viewer = createPptxViewer(document.getElementById('deck')!, {
	source: bytes,
	editable: true,
	customization,
});

document.getElementById('hide-insert')!.addEventListener('click', () => {
	viewer.hideRibbonTab('insert');
});
```

:::

## 用你自己的界面替换区域 {#replace-a-region-with-your-own-ui}

隐藏是移除一个区域；替换则保留其在布局中的位置，并把这个位置交给你。查看器会在内置界面原本所在之处（检查器所在的列、“选项”对话框的主体、“文件”选项卡的某个页面）渲染一个空容器，由你用任意框架或不用框架来填充。将渲染器放在 `regions` 下：

```ts
const customization: ViewerCustomization = {
	regions: {
		'dialog.options': 'acme-settings', // 你注册的自定义元素
		inspector: (container, context) => {
			const view = new AcmeInspector(container, context.viewer);
			view.show(context.selectedElementIds);
			return {
				update: (next) => view.show(next.selectedElementIds),
				destroy: () => view.dispose(),
			};
		},
	},
};
```

### 区域 {#regions}

| 区域                                               | 你替换的内容                                                          | 查看器保留的内容                                    |
| -------------------------------------------------- | --------------------------------------------------------------------- | --------------------------------------------------- |
| `inspector`                                        | 检查器窗格                                                            | 其所在列、宽度、调整手柄、打开/关闭状态和移动端面板 |
| `slidesPane`                                       | 幻灯片缩略图栏                                                        | 其宽度和折叠状态                                    |
| `notes`                                            | 演讲者备注面板                                                        | 其高度、调整手柄和折叠状态                          |
| `statusBar`                                        | 状态栏                                                                | 其在画布下方的行                                    |
| `dialog.options`                                   | “文件 > 选项”（设置对话框）的主体                                     | 对话框外框：背景遮罩、标题、关闭按钮、Escape        |
| `dialog.share`, `dialog.broadcast`, `dialog.print` | 该对话框的主体                                                        | 对话框外框                                          |
| `backstage.<page>`                                 | “文件”选项卡某个页面的主体（`backstage.info`、`backstage.export` 等） | “文件”选项卡导航、标题和返回按钮                    |

完整列表见[参考](#replaceable-regions-regions)。隐藏优先于替换：同时被隐藏的区域（通过 `hiddenPanels`、`hiddenDialogs`、`backstage.hiddenPages` 或某项能力）根本不会被渲染。

### 渲染器形式 {#renderer-forms}

区域接受三种形式中的任意一种。选择适合你技术栈的即可：

- **挂载函数** `(container, context) => cleanup | handle | void`。它在区域出现时运行一次。返回一个清理函数，或返回句柄 `{ update?(context), destroy?() }` 以同时接收上下文变化。
- **自定义元素标签**，以字符串形式给出（`'acme-settings'`）。查看器创建该元素，将其追加到容器中并赋值其 `context` 属性，此后每当上下文变化都会再次赋值 `context`。请在区域首次出现之前自行注册该元素。
- **自定义元素规格** `{ element: 'acme-settings', properties: { tenant } }`。与上面相同，只是 `properties` 会在元素挂接之前赋值一次。再次传入相等的规格（相同标签、相同属性值）会保留该元素。

只有当渲染器改变时，查看器才会重新挂载区域：不同的函数、不同的标签或不同的 `properties`。请只定义一次挂载函数（放在渲染函数之外，或使用 memo 缓存），就像对 `customization` 对象的其余部分那样。

抛出异常的渲染器会被捕获并记录；它绝不会使查看器崩溃。

### 区域何时挂载 {#when-a-region-is-mounted}

被替换的区域会在内置界面出现的位置和时机出现，并且查看器保留该槽位自身的行为：

- **可见性规则沿用。** 只读的演示文稿没有检查器，因此也没有宿主检查器；内置界面仅在桌面端显示或仅在编辑时显示的区域，遵循同样的规则。
- **折叠是隐藏，而不是卸载。** 折叠的备注或关闭的检查器会让你的渲染器保持挂载并隐藏，因此其状态得以保留；只有当该区域不再被替换、其渲染器改变或查看器被销毁时，它才会被销毁。
- **尺寸仍由查看器管理。** 检查器保留其宽度和调整手柄，备注面板保留其高度、调整手柄和折叠开关。你的内容填满容器（`width` 和 `height` 均为 100%）。
- **手机布局**会在容纳检查器、备注和幻灯片栏的底部面板内替换相同的内容。
- 对话框上的 **`close()`** 行为与对话框的关闭（X）按钮一致：它会关闭对话框，并保留通过 `setOption` 所做的所有更改。

每个容器都带有 `data-pptx-region="<id>"` 属性，这是你自己的 CSS 和测试的稳定挂钩。

### 上下文 {#the-context}

每个渲染器都会收到一个 `ViewerRegionContext`：

| 字段                             | 含义                                                                                                                                   |
| -------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| `region`                         | 区域 id（`inspector`、`dialog.options`、`backstage.host.<id>` 等）                                                                     |
| `viewer`                         | 查看器的命令式 API，与组件句柄是同一个对象：`getSlides()`、`getElementById()`、`updateElement()`、`goTo()`、`undo()`、自定义辅助方法等 |
| `mode`                           | `'edit'`、`'preview'`、`'present'` 或 `'master'`                                                                                       |
| `canEdit`                        | 演示文稿只读时为 `false`（edit 属性、`capabilities.edit` 或保护）                                                                      |
| `activeSlideIndex`, `slideCount` | 当前幻灯片（从零开始）和幻灯片数量                                                                                                     |
| `selectedElementIds`             | 当前幻灯片上选中的元素（为空：未选中任何内容）                                                                                         |
| `locale`                         | 当前界面区域设置                                                                                                                       |
| `close()`                        | 关闭对话框、离开“文件”选项卡或关闭检查器；对状态栏、备注和幻灯片窗格不执行任何操作                                                     |

`dialog.options` 还会添加“文件 > 选项”模型，使宿主设置视图能够读取和写入与内置对话框相同的设置：

| 字段                   | 含义                                                                       |
| ---------------------- | -------------------------------------------------------------------------- |
| `options`              | 当前的 `ViewerOptions`（已应用宿主锁定）                                   |
| `setOption(id, value)` | 按 `<group>.<key>` id（`advanced.showGrid`）更改一个设置；锁定的设置会忽略 |
| `resetOptions()`       | 恢复为宿主默认值或内置默认值                                               |

`close`、`setOption` 和 `resetOptions` 是稳定的：保留旧的上下文并在稍后调用它们是安全的。

每个绑定都在共享的 `<pptx-ui-options-dialog>` 元素中显示“文件 > 选项”。被替换的主体会被放入其槽位（`slot="body"`），位于页面自己的 DOM 中而不是该元素的 shadow 树里，因此你的页面样式可以作用到它。你的视图自行确定尺寸；当它比窗口高时，对话框会让它滚动。关闭按钮和 Escape 会保留你通过 `setOption` 所做的修改，与内置对话框相同。

### 挂载框架组件 {#mounting-a-framework-component}

挂载函数与框架无关，因此区域可以容纳任何框架的组件，包括与查看器绑定不同的框架：

::: code-group

```tsx [React]
import type { ViewerRegionMount } from 'ooxml-ui/pptx';
import { createRoot } from 'react-dom/client';

const settings: ViewerRegionMount = (container, context) => {
	const root = createRoot(container);
	root.render(<AcmeSettings context={context} />);
	return {
		update: (next) => root.render(<AcmeSettings context={next} />),
		destroy: () => root.unmount(),
	};
};

<PowerPointViewer content={bytes} customization={{ regions: { 'dialog.options': settings } }} />;
```

```ts [Vue]
import type { ViewerRegionContext, ViewerRegionMount } from 'ooxml-ui/pptx';
import { createApp, shallowRef } from 'vue';

const settings: ViewerRegionMount = (container, context) => {
	const current = shallowRef<ViewerRegionContext>(context);
	const app = createApp(AcmeSettings, { context: current });
	app.mount(container);
	return { update: (next) => (current.value = next), destroy: () => app.unmount() };
};
```

```ts [Angular]
import { createComponent, type ApplicationRef, type EnvironmentInjector } from '@angular/core';
import type { ViewerRegionMount } from 'pptx-angular-viewer';

function settingsRegion(app: ApplicationRef, injector: EnvironmentInjector): ViewerRegionMount {
	return (container, context) => {
		const ref = createComponent(AcmeSettingsComponent, {
			environmentInjector: injector,
			hostElement: container.appendChild(document.createElement('acme-settings-host')),
		});
		ref.setInput('context', context);
		app.attachView(ref.hostView);
		return {
			update: (next) => ref.setInput('context', next),
			destroy: () => {
				app.detachView(ref.hostView);
				ref.destroy();
			},
		};
	};
}
```

```ts [Svelte]
import type { ViewerRegionMount } from 'ooxml-ui/pptx';
// settings-region.svelte.ts（runes 需要 .svelte.ts 模块）
import { mount, unmount } from 'svelte';

const settings: ViewerRegionMount = (container, context) => {
	const props = $state({ context });
	const component = mount(AcmeSettings, { target: container, props });
	return { update: (next) => (props.context = next), destroy: () => unmount(component) };
};
```

```ts [Vanilla]
import type { ViewerRegionMount, ViewerSettingsRegionContext } from 'ooxml-ui/pptx';

const settings: ViewerRegionMount = (container, context) => {
	// `dialog.options` 在基础上下文之上还会收到设置模型。
	const { options, setOption } = context as ViewerSettingsRegionContext;
	const grid = document.createElement('label');
	const box = document.createElement('input');
	box.type = 'checkbox';
	box.checked = options.advanced.showGrid;
	box.addEventListener('change', () => setOption('advanced.showGrid', box.checked));
	grid.append(box, ' Show gridlines');
	container.append(grid);
};
```

:::

### 你自己的“文件”选项卡页面 {#your-own-file-tab-pages}

`backstage.extraPages` 会在“文件”选项卡中其所属分组的内置页面之后添加页面。每个页面的主体是一个区域（`backstage.host.<id>`），接受相同的渲染器形式：

```ts
const customization: ViewerCustomization = {
	backstage: {
		extraPages: [
			{ id: 'policies', label: 'Policies', render: 'acme-policies' },
			{ id: 'about', label: 'About Acme', group: 'footer', render: mountAbout },
		],
	},
};
```

`label` 是你自己的文本，不会被翻译。页面 id 在你的页面之间必须唯一；重复的 id 会被忽略。`addBackstagePage(page)` 和 `removeBackstagePage(id)` 可在运行时更改它们。

## 能力 {#capabilities}

`capabilities` 声明用户可以做什么，每项能力一个开关，并移除通向被关闭内容的所有路径：按钮、菜单项、快捷键、“文件”页面和卡片。省略某个键（或为 `true`）则由查看器自身的规则决定。

```ts
const portal: ViewerCustomization = {
	capabilities: { edit: false, export: false, print: false, copy: false },
};
```

| 能力      | `false` 的含义                                                          |
| --------- | ----------------------------------------------------------------------- |
| `edit`    | 演示文稿以只读方式打开，无论编辑属性（`canEdit`、`editable`）如何设置。 |
| `export`  | 没有“文件 > 导出”页面、导出按钮或导出卡片。                             |
| `print`   | 没有打印对话框、“文件 > 打印”页面或打印卡片。                           |
| `share`   | 没有共享或广播对话框、按钮、页面或卡片。                                |
| `copy`    | 没有复制、剪切或另存为图片（菜单、功能区、快捷键），也没有复制为图像。  |
| `present` | 没有“幻灯片放映”标签页或 F5 / Shift+F5（`presentMode` 功能）。          |
| `comment` | 没有“添加批注”项，“插入”和“审阅”中也没有批注组（`comments` 功能）。     |

能力是界面策略，而不是文档保护。使用浏览器开发者工具的用户仍然拥有你交给查看器的字节；如果某个演示文稿不得被复制或导出，就不要把它发送到浏览器。

`setCapability(id, allowed)` 可在运行时更改其中一项；`undefined` 恢复默认值。

## 示例 {#recipes}

### 展台：没有编辑界面的查看器 {#kiosk-a-viewer-with-no-editing-chrome}

只读已经隐藏了编辑命令；这里还会移除展台访客可能误入的所有内容。

```ts
const kiosk: ViewerCustomization = {
	ribbon: {
		hiddenTabs: [
			'file',
			'home',
			'insert',
			'draw',
			'design',
			'transitions',
			'animations',
			'record',
			'review',
			'view',
			'help',
		],
		hiddenButtons: ['share', 'broadcast', 'export', 'undo', 'redo', 'record'],
	},
	hiddenPanels: ['inspector', 'notes', 'quickAccessToolbar'],
	disabledFeatures: ['ai', 'collaboration', 'comments'],
	hiddenDialogs: ['options', 'print', 'export'],
	contextMenu: { disableElementMenu: true, disableCanvasMenu: true },
	keyboard: { disableAll: true },
};
```

保留 `slideShow`（以及 `navigation` 和 `fullscreen` 按钮）可见，访客仍可放映。

### 隐藏用户不需要的设置页面 {#hide-the-settings-pages-users-do-not-need}

```ts
const settings: ViewerCustomization = {
	options: {
		hiddenPages: ['proofing', 'addIns', 'trust', 'quickAccess', 'ribbon'],
		hiddenSections: ['general.startup', 'save.cache', 'advanced.print'],
		hiddenSettings: ['advanced.disableHardwareAcceleration'],
	},
};
```

要完全移除“文件 > 选项”，请使用 `hiddenDialogs: ['options']`。

### 固定用户身份、区域设置和主题 {#pin-the-user-identity-locale-and-theme}

作者名来自已登录用户，不得被编辑：

```ts
viewer.lockSetting('general.userName', currentUser.displayName);
viewer.lockSetting('general.userInitials', currentUser.initials, true); // 锁定并隐藏
```

区域设置和主题是宿主属性，而不是“选项”中的设置。请通过你所用绑定的区域设置和主题属性来固定它们（`defaultLocale` / `locale`、`theme` / `defaultThemeKey`，参见[本地化](/zh/guide/localization)和[主题](/zh/guide/theming)指南），然后移除选择器：

```ts
const pinned: ViewerCustomization = {
	options: {
		hiddenPages: ['language'], // 显示语言选择器
		hiddenSections: ['general.appearance'], // 查看器主题选择器
	},
};
```

### 移除 AI 和协作 {#remove-ai-and-collaboration}

```ts
const noCloud: ViewerCustomization = { disabledFeatures: ['ai', 'collaboration'] };
```

禁用 `ai` 后，即使传入了 `ai` 配置，助手也保持关闭，这样同一个构建就可以发布给有和没有该附加组件的租户：

```ts
viewer.setFeatureEnabled('ai', tenant.hasAiAddOn);
```

### 将功能区精简为最小集合 {#trim-the-ribbon-to-a-minimal-set}

```ts
import { RIBBON_TAB_IDS } from 'pptx-react-viewer';

const keep = new Set(['file', 'home', 'insert']);
const minimal: ViewerCustomization = {
	ribbon: {
		hiddenTabs: RIBBON_TAB_IDS.filter((id) => !keep.has(id)),
		hiddenButtons: ['broadcast', 'record'],
	},
};
```

### 公司快捷键 {#company-shortcuts}

```ts
const keys: ViewerCustomization = {
	keyboard: {
		disabled: ['newSlide', 'toggleShortcuts'],
		remap: { duplicate: 'Mod+Shift+D', group: ['Mod+G', 'Alt+G'] },
	},
};
```

### 你自己的设置视图 {#your-own-settings-view}

保留查看器“选项”对话框的外框，但显示你自己的设置，读写同一个模型：

```ts
customElements.define(
	'acme-settings',
	class extends HTMLElement {
		set context(context: ViewerSettingsRegionContext) {
			this.textContent = '';
			const grid = document.createElement('input');
			grid.type = 'checkbox';
			grid.checked = context.options.advanced.showGrid;
			grid.onchange = () => context.setOption('advanced.showGrid', grid.checked);
			const done = document.createElement('button');
			done.textContent = 'Done';
			done.onclick = () => context.close();
			this.append(grid, done);
		}
	},
);

const customization: ViewerCustomization = { regions: { 'dialog.options': 'acme-settings' } };
```

若要改为将用户引导到你应用自己的设置页面，请隐藏该对话框（`hiddenDialogs: ['options']`），并添加一个导航到那里的“文件”选项卡页面或功能区加载项命令。

### 只读门户 {#a-read-only-portal}

```ts
const portal: ViewerCustomization = {
	capabilities: { edit: false, export: false, print: false, share: false, copy: false },
	disabledFeatures: ['ai'],
	hiddenPanels: ['inspector'],
};
```

### 构建你自己的管理界面 {#build-your-own-admin-screen}

每个 id 列表都与类型一起以运行时数组的形式导出：`RIBBON_TAB_IDS`、`TOOLBAR_BUTTON_IDS`、`OPTIONS_PAGE_IDS`、`OPTIONS_SECTION_IDS`、`OPTIONS_SETTING_IDS`、`BACKSTAGE_PAGE_IDS`、`BACKSTAGE_CARD_IDS`、`ELEMENT_CONTEXT_MENU_COMMAND_IDS`、`CANVAS_CONTEXT_MENU_COMMAND_IDS`、`EDITOR_SHORTCUT_ACTION_IDS`、`VIEWER_PANEL_IDS`、`VIEWER_FEATURE_IDS`、`VIEWER_DIALOG_IDS`、`VIEWER_EXPORT_FORMAT_IDS`。`VIEWER_CAPABILITY_IDS` 和 `VIEWER_REGION_IDS` 从 `ooxml-ui/pptx` 导出。用它们渲染复选框，按租户保存得到的对象，并通过 `setCustomization` 回传。

## 目前尚不可自定义的内容 {#what-is-not-customisable-yet}

- **下拉菜单或库内的控件**（项目符号库中的单个条目、一个形状样式图块）无法单独寻址；请改为隐藏打开该菜单的控件。
- **放映、幻灯片浏览和演示者视图的键盘映射。** `keyboard` 只涵盖编辑器键盘映射。它唯一影响的放映键是 F5 / Shift+F5，通过 `presentMode` 功能实现。
- **键盘快捷键参考**（`?` 浮层和“选项 > 自定义功能区 > 键盘快捷键”）列出的是内置组合键，而不是你的重映射。如果这很重要，请用 `keyboard.disabled: ['toggleShortcuts']` 和 `options.hiddenSections: ['ribbon.shortcutReference']` 将其隐藏。
- **检查器分区**和幻灯片栏自身的右键菜单无法单独寻址；请用 `hiddenPanels` 隐藏整个区域，或用 `regions.inspector` 替换整个检查器。
- **可替换的区域**是[区域](#regions)下列出的那些。功能区、标题栏、快速访问栏和其他对话框可以隐藏但不能替换；请改用 `ribbonAddIns` 添加你自己的功能区标签页。
- **框架原生插槽**（React render props、Vue 插槽、Angular 模板、Svelte 代码片段）不予提供；每个绑定都接受相同的挂载函数或自定义元素，它们可以渲染任何框架的组件（参见[挂载框架组件](#mounting-a-framework-component)）。
- **移动端布局**遵循功能区、菜单、对话框和功能的自定义，但移动端底部面板没有自己的面板 id，并且某些移动端底栏按钮（幻灯片、备注、格式）在其面板被隐藏后仍然可见（点按它们不会有任何反应）。
- **用户添加的快速访问命令**（例如打印）在其对话框被移除后仍留在栏中；它们是失效而非隐藏。
- **`presentMode`** 移除“幻灯片放映”标签页和 F5 / Shift+F5，但不会移除所有次要的“放映”入口（例如某些绑定中视图切换器的放映选项）。
- **Vanilla** 没有 `update(options)`，因此它的 `customization` 选项只是初始值；请使用 `setCustomization` 和其他辅助方法在运行时更改。更改会重建界面（与 `setLocale` 相同），这会使功能区回到第一个标签页。
- **React 的无头 `useViewerBuildingBlocks`** 组合不接受 `customization`；请使用 `PowerPointViewer` 组件。

## 参考 {#reference}

由共享的 id 目录生成（`bun run docs:customization`）；如果此处缺少某个 id，单元测试会失败。

<!-- customization-reference:start -->

### Ribbon tabs (`ribbon.hiddenTabs`)

| Id               | Tab             |
| ---------------- | --------------- |
| `file`           | File            |
| `home`           | Home            |
| `insert`         | Insert          |
| `draw`           | Draw            |
| `design`         | Design          |
| `transitions`    | Transitions     |
| `animations`     | Animations      |
| `slideShow`      | Slide Show      |
| `record`         | Record          |
| `review`         | Review          |
| `view`           | View            |
| `help`           | Help            |
| `shapeFormat`    | Shape Format    |
| `pictureFormat`  | Picture Format  |
| `tableDesign`    | Table Design    |
| `tableLayout`    | Table Layout    |
| `chartDesign`    | Chart Design    |
| `chartFormat`    | Chart Format    |
| `smartArtDesign` | SmartArt Design |

### Ribbon groups (`ribbon.hiddenGroups`)

| Id                                  | Group                    |
| ----------------------------------- | ------------------------ |
| `home.clipboard`                    | Clipboard                |
| `home.slides`                       | Slides                   |
| `home.font`                         | Font                     |
| `home.paragraph`                    | Paragraph                |
| `home.drawing`                      | Drawing                  |
| `home.arrange`                      | Arrange                  |
| `home.editing`                      | Editing                  |
| `insert.slides`                     | Slides                   |
| `insert.tables`                     | Tables                   |
| `insert.images`                     | Images                   |
| `insert.camera`                     | Camera                   |
| `insert.illustrations`              | Illustrations            |
| `insert.forms`                      | Forms                    |
| `insert.addIns`                     | Add-ins                  |
| `insert.links`                      | Links                    |
| `insert.comments`                   | Comments                 |
| `insert.text`                       | Text                     |
| `insert.symbols`                    | Symbols                  |
| `insert.media`                      | Media                    |
| `draw.tools`                        | Drawing Tools            |
| `draw.convert`                      | Convert                  |
| `design.themes`                     | Themes                   |
| `design.variants`                   | Variants                 |
| `design.customize`                  | Customize                |
| `design.designer`                   | Designer                 |
| `transitions.preview`               | Preview                  |
| `transitions.transitionToThisSlide` | Transition to This Slide |
| `transitions.timing`                | Timing                   |
| `animations.preview`                | Preview                  |
| `animations.animation`              | Animation                |
| `animations.motionPath`             | Motion Paths             |
| `animations.advancedAnimation`      | Advanced Animation       |
| `animations.timing`                 | Timing                   |
| `slideShow.startSlideShow`          | Start Slide Show         |
| `slideShow.present`                 | Present                  |
| `slideShow.setUp`                   | Set Up                   |
| `slideShow.captions`                | Captions & Subtitles     |
| `record.camera`                     | Camera                   |
| `record.record`                     | Record                   |
| `record.manage`                     | Manage                   |
| `record.help`                       | Help                     |
| `review.proofing`                   | Proofing                 |
| `review.accessibility`              | Accessibility            |
| `review.language`                   | Language                 |
| `review.comments`                   | Comments                 |
| `review.compare`                    | Compare                  |
| `review.protect`                    | Protect                  |
| `review.ink`                        | Ink                      |
| `view.presentationViews`            | Presentation Views       |
| `view.masterViews`                  | Master Views             |
| `view.show`                         | Show                     |
| `view.zoom`                         | Zoom                     |
| `view.window`                       | Window                   |
| `help.help`                         | Help                     |
| `shapeFormat.shapeStyles`           | Shape Styles             |
| `shapeFormat.wordArtStyles`         | WordArt Styles           |
| `pictureFormat.adjust`              | Adjust                   |
| `pictureFormat.pictureStyles`       | Picture Styles           |
| `pictureFormat.size`                | Size                     |
| `tableDesign.tableStyles`           | Table Styles             |
| `tableLayout.table`                 | Table                    |
| `tableLayout.rowsColumns`           | Rows & Columns           |
| `tableLayout.merge`                 | Merge                    |
| `tableLayout.cellSize`              | Cell Size                |
| `tableLayout.alignment`             | Alignment                |
| `tableLayout.tableSize`             | Table Size               |
| `tableLayout.arrange`               | Arrange                  |
| `chartDesign.chartLayouts`          | Chart Layouts            |
| `chartDesign.chartStyles`           | Chart Styles             |
| `chartFormat.currentSelection`      | Current Selection        |
| `smartArtDesign.createGraphic`      | Create Graphic           |
| `smartArtDesign.layouts`            | Layouts                  |
| `smartArtDesign.smartArtStyles`     | SmartArt Styles          |
| `smartArtDesign.reset`              | Reset                    |

### Ribbon controls (`ribbon.hiddenButtons`)

| Id                                                | Control                                                                       |
| ------------------------------------------------- | ----------------------------------------------------------------------------- |
| `home.clipboard.paste`                            | Paste                                                                         |
| `home.clipboard.cut`                              | Cut                                                                           |
| `home.clipboard.copy`                             | Copy                                                                          |
| `home.clipboard.formatPainter`                    | Format Painter                                                                |
| `home.slides.newSlide`                            | New Slide                                                                     |
| `home.slides.layout`                              | Layout                                                                        |
| `home.slides.reset`                               | Reset                                                                         |
| `home.slides.section`                             | Section                                                                       |
| `home.slides.slideTemplates`                      | Slide templates (New Slide menu entry)                                        |
| `home.font.fontFamily`                            | Font                                                                          |
| `home.font.fontSize`                              | Font Size                                                                     |
| `home.font.increaseFontSize`                      | Increase Font Size                                                            |
| `home.font.decreaseFontSize`                      | Decrease Font Size                                                            |
| `home.font.clearFormatting`                       | Clear All Formatting                                                          |
| `home.font.bold`                                  | Bold                                                                          |
| `home.font.italic`                                | Italic                                                                        |
| `home.font.underline`                             | Underline                                                                     |
| `home.font.strikethrough`                         | Strikethrough                                                                 |
| `home.font.shadow`                                | Text Shadow                                                                   |
| `home.font.characterSpacing`                      | Character Spacing                                                             |
| `home.font.changeCase`                            | Change Case                                                                   |
| `home.font.fontColor`                             | Font Color                                                                    |
| `home.font.highlightColor`                        | Text Highlight Color                                                          |
| `home.font.superscript`                           | Superscript                                                                   |
| `home.font.subscript`                             | Subscript                                                                     |
| `home.paragraph.bullets`                          | Bullets (toggle and gallery)                                                  |
| `home.paragraph.numbering`                        | Numbering (toggle and gallery)                                                |
| `home.paragraph.decreaseIndent`                   | Decrease List Level                                                           |
| `home.paragraph.increaseIndent`                   | Increase List Level                                                           |
| `home.paragraph.lineSpacing`                      | Line Spacing                                                                  |
| `home.paragraph.alignLeft`                        | Align Left                                                                    |
| `home.paragraph.alignCenter`                      | Center                                                                        |
| `home.paragraph.alignRight`                       | Align Right                                                                   |
| `home.paragraph.justify`                          | Justify                                                                       |
| `home.paragraph.columns`                          | Columns                                                                       |
| `home.paragraph.textDirection`                    | Text Direction                                                                |
| `home.paragraph.alignText`                        | Align Text                                                                    |
| `home.drawing.shapes`                             | Shapes                                                                        |
| `home.drawing.arrange`                            | Arrange                                                                       |
| `home.drawing.quickStyles`                        | Quick Styles (Shape Styles gallery)                                           |
| `home.drawing.shapeFill`                          | Shape Fill                                                                    |
| `home.drawing.shapeOutline`                       | Shape Outline                                                                 |
| `home.drawing.shapeEffects`                       | Shape Effects gallery                                                         |
| `home.arrange.bringForward`                       | Bring Forward                                                                 |
| `home.arrange.sendBackward`                       | Send Backward                                                                 |
| `home.arrange.bringToFront`                       | Bring to Front                                                                |
| `home.arrange.sendToBack`                         | Send to Back                                                                  |
| `home.arrange.flipHorizontal`                     | Flip Horizontal                                                               |
| `home.arrange.flipVertical`                       | Flip Vertical                                                                 |
| `home.arrange.duplicate`                          | Duplicate                                                                     |
| `home.arrange.delete`                             | Delete                                                                        |
| `home.arrange.group`                              | Group                                                                         |
| `home.arrange.ungroup`                            | Ungroup                                                                       |
| `home.arrange.align`                              | Align                                                                         |
| `home.arrange.mergeShapes`                        | Merge Shapes                                                                  |
| `home.arrange.crop`                               | Crop                                                                          |
| `home.arrange.outlineWidth`                       | Outline width                                                                 |
| `home.editing.find`                               | Find                                                                          |
| `home.editing.replace`                            | Replace                                                                       |
| `home.editing.select`                             | Select                                                                        |
| `insert.slides.newSlide`                          | New Slide                                                                     |
| `insert.slides.reuseSlides`                       | Reuse Slides                                                                  |
| `insert.tables.table`                             | Table                                                                         |
| `insert.images.pictures`                          | Pictures                                                                      |
| `insert.images.screenshot`                        | Screenshot                                                                    |
| `insert.images.photoAlbum`                        | Photo Album                                                                   |
| `insert.camera.cameo`                             | Cameo                                                                         |
| `insert.illustrations.shapes`                     | Shapes                                                                        |
| `insert.illustrations.icons`                      | Icons                                                                         |
| `insert.illustrations.threeDModels`               | 3D Models                                                                     |
| `insert.illustrations.smartArt`                   | SmartArt                                                                      |
| `insert.illustrations.chart`                      | Chart                                                                         |
| `insert.forms.forms`                              | Forms                                                                         |
| `insert.addIns.getAddIns`                         | Get Add-ins                                                                   |
| `insert.addIns.myAddIns`                          | My Add-ins                                                                    |
| `insert.links.zoom`                               | Zoom                                                                          |
| `insert.links.link`                               | Link                                                                          |
| `insert.links.action`                             | Action button                                                                 |
| `insert.comments.comment`                         | Comment                                                                       |
| `insert.text.textBox`                             | Text Box                                                                      |
| `insert.text.headerFooter`                        | Header & Footer                                                               |
| `insert.text.wordArt`                             | WordArt                                                                       |
| `insert.text.dateTime`                            | Date & Time                                                                   |
| `insert.text.slideNumber`                         | Slide Number                                                                  |
| `insert.text.object`                              | Object                                                                        |
| `insert.text.field`                               | Date & Time and Slide Number (deprecated alias: use dateTime and slideNumber) |
| `insert.symbols.equation`                         | Equation                                                                      |
| `insert.symbols.symbol`                           | Symbol                                                                        |
| `insert.media.media`                              | Video                                                                         |
| `insert.media.audio`                              | Audio                                                                         |
| `insert.media.screenRecording`                    | Screen Recording                                                              |
| `draw.tools.select`                               | Select                                                                        |
| `draw.tools.pen`                                  | Pen                                                                           |
| `draw.tools.highlighter`                          | Highlighter                                                                   |
| `draw.tools.freeform`                             | Freeform                                                                      |
| `draw.tools.eraser`                               | Eraser                                                                        |
| `draw.tools.penColor`                             | Pen colour                                                                    |
| `draw.tools.penWidth`                             | Pen width                                                                     |
| `draw.convert.inkToShape`                         | Ink to Shape                                                                  |
| `design.themes.gallery`                           | Themes (inline gallery)                                                       |
| `design.themes.browseThemes`                      | Themes gallery                                                                |
| `design.themes.editTheme`                         | Edit theme                                                                    |
| `design.variants.gallery`                         | Variants (inline gallery)                                                     |
| `design.variants.colors`                          | Variants: Colors gallery                                                      |
| `design.variants.fonts`                           | Variants: Fonts gallery                                                       |
| `design.variants.effects`                         | Variants: Effects gallery                                                     |
| `design.variants.backgroundStyles`                | Variants: Background Styles gallery                                           |
| `design.customize.slideSize`                      | Slide Size                                                                    |
| `design.customize.formatBackground`               | Format Background                                                             |
| `design.designer.designIdeas`                     | Design Ideas                                                                  |
| `transitions.preview.preview`                     | Preview                                                                       |
| `transitions.transitionToThisSlide.gallery`       | Transition gallery                                                            |
| `transitions.transitionToThisSlide.effectOptions` | Effect Options                                                                |
| `transitions.timing.sound`                        | Sound                                                                         |
| `transitions.timing.duration`                     | Duration                                                                      |
| `transitions.timing.applyToAll`                   | Apply To All                                                                  |
| `transitions.timing.advanceOnClick`               | On Mouse Click                                                                |
| `transitions.timing.advanceAfter`                 | After                                                                         |
| `animations.preview.preview`                      | Preview                                                                       |
| `animations.animation.gallery`                    | Animation gallery                                                             |
| `animations.animation.effectOptions`              | Effect Options                                                                |
| `animations.motionPath.gallery`                   | Motion path gallery                                                           |
| `animations.advancedAnimation.addAnimation`       | Add Animation                                                                 |
| `animations.advancedAnimation.animationPane`      | Animation Pane                                                                |
| `animations.advancedAnimation.trigger`            | Trigger                                                                       |
| `animations.advancedAnimation.animationPainter`   | Animation Painter                                                             |
| `animations.advancedAnimation.remove`             | Remove animation                                                              |
| `animations.timing.start`                         | Start                                                                         |
| `animations.timing.duration`                      | Duration                                                                      |
| `animations.timing.delay`                         | Delay                                                                         |
| `animations.timing.reorder`                       | Reorder                                                                       |
| `slideShow.startSlideShow.fromBeginning`          | From Beginning                                                                |
| `slideShow.startSlideShow.fromCurrent`            | From Current Slide                                                            |
| `slideShow.startSlideShow.customShow`             | Custom Slide Show                                                             |
| `slideShow.present.presenterView`                 | Presenter View                                                                |
| `slideShow.present.broadcast`                     | Present Online                                                                |
| `slideShow.setUp.setUpSlideShow`                  | Set Up Slide Show                                                             |
| `slideShow.setUp.hideSlide`                       | Hide Slide                                                                    |
| `slideShow.setUp.rehearseTimings`                 | Rehearse Timings                                                              |
| `slideShow.setUp.record`                          | Record                                                                        |
| `slideShow.setUp.rehearseWithCoach`               | Rehearse with Coach                                                           |
| `slideShow.captions.subtitles`                    | Always Use Subtitles                                                          |
| `slideShow.captions.subtitleSettings`             | Subtitle Settings                                                             |
| `record.camera.cameo`                             | Cameo                                                                         |
| `record.record.fromBeginning`                     | From Beginning                                                                |
| `record.record.fromCurrent`                       | From Current Slide                                                            |
| `record.manage.clear`                             | Clear                                                                         |
| `record.manage.reset`                             | Reset to Cameo                                                                |
| `record.help.learnMore`                           | Learn more                                                                    |
| `review.proofing.spelling`                        | Spelling                                                                      |
| `review.proofing.thesaurus`                       | Thesaurus                                                                     |
| `review.accessibility.check`                      | Check Accessibility                                                           |
| `review.language.translate`                       | Translate                                                                     |
| `review.language.language`                        | Language                                                                      |
| `review.comments.newComment`                      | New Comment                                                                   |
| `review.comments.delete`                          | Delete                                                                        |
| `review.comments.deleteSlide`                     | Delete Slide Comments                                                         |
| `review.comments.deleteAll`                       | Delete All Comments                                                           |
| `review.comments.previous`                        | Previous                                                                      |
| `review.comments.next`                            | Next                                                                          |
| `review.comments.showComments`                    | Show Comments                                                                 |
| `review.compare.compare`                          | Compare                                                                       |
| `review.compare.markAllRead`                      | Mark all read                                                                 |
| `review.protect.readOnly`                         | Read-only                                                                     |
| `review.protect.restrictPermission`               | Restrict Permission                                                           |
| `review.ink.hideInk`                              | Hide Ink                                                                      |
| `view.presentationViews.normal`                   | Normal                                                                        |
| `view.presentationViews.outline`                  | Outline View                                                                  |
| `view.presentationViews.slideSorter`              | Slide Sorter                                                                  |
| `view.presentationViews.notesPage`                | Notes Page                                                                    |
| `view.presentationViews.readingView`              | Reading View                                                                  |
| `view.masterViews.slideMaster`                    | Slide Master                                                                  |
| `view.masterViews.handoutMaster`                  | Handout Master                                                                |
| `view.masterViews.notesMaster`                    | Notes Master                                                                  |
| `view.show.ruler`                                 | Ruler                                                                         |
| `view.show.gridlines`                             | Gridlines                                                                     |
| `view.show.guides`                                | Guides                                                                        |
| `view.show.snapToGrid`                            | Snap to Grid                                                                  |
| `view.show.snapToShape`                           | Snap to Shape                                                                 |
| `view.show.addGuide`                              | Add horizontal / vertical guide                                               |
| `view.show.selectionPane`                         | Selection Pane                                                                |
| `view.show.eyedropper`                            | Eyedropper                                                                    |
| `view.show.notes`                                 | Notes                                                                         |
| `view.zoom.zoom`                                  | Zoom                                                                          |
| `view.zoom.fitToWindow`                           | Fit to Window                                                                 |
| `view.window.templateEditing`                     | Edit template elements                                                        |
| `view.window.macros`                              | Macros                                                                        |
| `help.help.options`                               | Options                                                                       |
| `help.help.keyboardShortcuts`                     | Keyboard shortcuts                                                            |
| `help.help.accessibility`                         | Accessibility checker                                                         |
| `shapeFormat.shapeStyles.gallery`                 | Shape Styles gallery                                                          |
| `shapeFormat.shapeStyles.shapeEffects`            | Shape Effects gallery                                                         |
| `shapeFormat.wordArtStyles.gallery`               | WordArt Styles gallery                                                        |
| `pictureFormat.adjust.corrections`                | Corrections gallery                                                           |
| `pictureFormat.adjust.color`                      | Color gallery                                                                 |
| `pictureFormat.adjust.artisticEffects`            | Artistic Effects gallery                                                      |
| `pictureFormat.pictureStyles.gallery`             | Picture Styles gallery                                                        |
| `pictureFormat.pictureStyles.pictureEffects`      | Picture Effects gallery                                                       |
| `pictureFormat.size.crop`                         | Crop                                                                          |
| `tableDesign.tableStyles.gallery`                 | Table Styles gallery                                                          |
| `tableLayout.table.select`                        | Select                                                                        |
| `tableLayout.table.viewGridlines`                 | View Gridlines                                                                |
| `tableLayout.rowsColumns.deleteRows`              | Delete Rows                                                                   |
| `tableLayout.rowsColumns.deleteColumns`           | Delete Columns                                                                |
| `tableLayout.rowsColumns.insertAbove`             | Insert Above                                                                  |
| `tableLayout.rowsColumns.insertBelow`             | Insert Below                                                                  |
| `tableLayout.rowsColumns.insertLeft`              | Insert Left                                                                   |
| `tableLayout.rowsColumns.insertRight`             | Insert Right                                                                  |
| `tableLayout.merge.mergeCells`                    | Merge Cells                                                                   |
| `tableLayout.merge.splitCells`                    | Split Cells                                                                   |
| `tableLayout.cellSize.distributeRows`             | Distribute Rows                                                               |
| `tableLayout.cellSize.distributeColumns`          | Distribute Columns                                                            |
| `tableLayout.alignment.alignLeft`                 | Align Left                                                                    |
| `tableLayout.alignment.alignCenter`               | Center                                                                        |
| `tableLayout.alignment.alignRight`                | Align Right                                                                   |
| `tableLayout.alignment.alignTop`                  | Align Top                                                                     |
| `tableLayout.alignment.alignMiddle`               | Center Vertically                                                             |
| `tableLayout.alignment.alignBottom`               | Align Bottom                                                                  |
| `tableLayout.alignment.textDirection`             | Text Direction                                                                |
| `tableLayout.alignment.cellMargins`               | Cell Margins                                                                  |
| `tableLayout.tableSize.size`                      | Table Size                                                                    |
| `tableLayout.arrange.arrange`                     | Arrange                                                                       |
| `chartDesign.chartLayouts.quickLayout`            | Quick Layout gallery                                                          |
| `chartDesign.chartStyles.changeColors`            | Change Colors gallery                                                         |
| `chartDesign.chartStyles.gallery`                 | Chart Styles gallery                                                          |
| `chartFormat.currentSelection.chartElements`      | Chart Elements                                                                |
| `chartFormat.currentSelection.formatSelection`    | Format Selection                                                              |
| `chartFormat.currentSelection.resetToMatchStyle`  | Reset to Match Style                                                          |
| `smartArtDesign.createGraphic.addShape`           | Add Shape                                                                     |
| `smartArtDesign.createGraphic.addShapeBefore`     | Add Shape Before                                                              |
| `smartArtDesign.createGraphic.addShapeAbove`      | Add Shape Above                                                               |
| `smartArtDesign.createGraphic.addBullet`          | Add Bullet                                                                    |
| `smartArtDesign.createGraphic.textPane`           | Text Pane                                                                     |
| `smartArtDesign.createGraphic.promote`            | Promote                                                                       |
| `smartArtDesign.createGraphic.demote`             | Demote                                                                        |
| `smartArtDesign.createGraphic.rightToLeft`        | Right to Left                                                                 |
| `smartArtDesign.createGraphic.moveUp`             | Move Up                                                                       |
| `smartArtDesign.createGraphic.moveDown`           | Move Down                                                                     |
| `smartArtDesign.layouts.gallery`                  | Layouts gallery                                                               |
| `smartArtDesign.smartArtStyles.changeColors`      | Change Colors gallery                                                         |
| `smartArtDesign.smartArtStyles.gallery`           | SmartArt Styles gallery                                                       |
| `smartArtDesign.reset.resetGraphic`               | Reset Graphic                                                                 |
| `smartArtDesign.reset.convert`                    | Convert                                                                       |

### Toolbar buttons (`ribbon.hiddenButtons`)

| Id            | What it removes                                                                                          |
| ------------- | -------------------------------------------------------------------------------------------------------- |
| `share`       | Share / collaboration button in the tab row and mobile toolbar.                                          |
| `broadcast`   | Broadcast (present online) button.                                                                       |
| `export`      | Export button and the File > Export page.                                                                |
| `undo`        | Undo button in the quick-access strip.                                                                   |
| `redo`        | Redo button in the quick-access strip.                                                                   |
| `record`      | Record button and the Record ribbon tab.                                                                 |
| `notes`       | Notes toggle in the status bar.                                                                          |
| `fullscreen`  | Full-screen toggle.                                                                                      |
| `zoom`        | The zoom cluster (zoom in, zoom out, fit).                                                               |
| `navigation`  | The previous / next slide cluster.                                                                       |
| `mergeShapes` | The Merge Shapes dropdown (Union, Combine, Fragment, Intersect, Subtract) in the Home tab Arrange group. |
| `crop`        | Both picture Crop buttons (crop mode, aspect ratio, Fill, Fit): Home > Arrange, Picture Format > Size.   |

### Options pages (`options.hiddenPages`)

| Id              | Page                 |
| --------------- | -------------------- |
| `general`       | General              |
| `proofing`      | Proofing             |
| `save`          | Save                 |
| `language`      | Language             |
| `accessibility` | Accessibility        |
| `advanced`      | Advanced             |
| `ribbon`        | Customize Ribbon     |
| `quickAccess`   | Quick Access Toolbar |
| `addIns`        | Add-ins              |
| `trust`         | Trust Center         |
| `ai`            | AI                   |

### Options sections (`options.hiddenSections`)

| Id                               | Section                                 |
| -------------------------------- | --------------------------------------- |
| `general.userInterface`          | User Interface options                  |
| `general.personalize`            | Personalize your copy of the viewer     |
| `general.appearance`             | Viewer theme                            |
| `general.fonts`                  | Fonts                                   |
| `general.startup`                | Start up options                        |
| `proofing.autoCorrect`           | AutoCorrect options                     |
| `proofing.spellingOffice`        | When correcting spelling                |
| `proofing.spellingViewer`        | When correcting spelling in the viewer  |
| `save.savePresentations`         | Save presentations                      |
| `save.cache`                     | Cache Settings                          |
| `accessibility.assistant`        | Make your document accessible to others |
| `accessibility.feedback`         | Feedback options                        |
| `accessibility.display`          | Application display options             |
| `advanced.editing`               | Editing options                         |
| `advanced.cutCopyPaste`          | Cut, copy, and paste                    |
| `advanced.imageQuality`          | Image Size and Quality                  |
| `advanced.chart`                 | Chart                                   |
| `advanced.display`               | Display                                 |
| `advanced.slideShow`             | Slide Show                              |
| `advanced.print`                 | Print                                   |
| `ribbon.shortcutReference`       | Keyboard Shortcuts                      |
| `quickAccess.quickAccessOptions` | Quick Access Toolbar options            |
| `trust.trustSettings`            | Security settings                       |

### Settings (`options.hiddenSettings`, `options.locked`, `options.defaults`)

| Id                                           | Type    | Built-in default | Label                                                           |
| -------------------------------------------- | ------- | ---------------- | --------------------------------------------------------------- |
| `general.displayOptimization`                | string  | `'appearance'`   | When using multiple displays                                    |
| `general.showMiniToolbar`                    | boolean | `true`           | Show Mini Toolbar on selection                                  |
| `general.enableLivePreview`                  | boolean | `true`           | Enable Live Preview                                             |
| `general.collapseRibbonAutomatically`        | boolean | `false`          | Collapse the ribbon automatically                               |
| `general.collapseSearchByDefault`            | boolean | `false`          | Collapse the search box by default                              |
| `general.screenTipStyle`                     | string  | `'descriptions'` | ScreenTip style                                                 |
| `general.userName`                           | string  | `''`             | User name                                                       |
| `general.userInitials`                       | string  | `''`             | Initials                                                        |
| `general.showStartScreen`                    | boolean | `true`           | Show the Start screen when this application starts              |
| `general.enableCustomFontUpload`             | boolean | `false`          | Let me add font files to this session                           |
| `proofing.autoCorrectTwoInitialCapitals`     | boolean | `true`           | Correct TWo INitial CApitals                                    |
| `proofing.autoCorrectCapitalizeFirstLetter`  | boolean | `true`           | Capitalize first letter of sentences                            |
| `proofing.autoCorrectCapitalizeDayNames`     | boolean | `true`           | Capitalize names of days                                        |
| `proofing.autoCorrectSmartQuotes`            | boolean | `true`           | Replace straight quotes with smart quotes                       |
| `proofing.autoCorrectHyphensToDash`          | boolean | `true`           | Replace hyphens (--) with dash                                  |
| `proofing.autoCorrectFractions`              | boolean | `true`           | Replace fractions (1/2) with fraction characters (½)            |
| `proofing.autoCorrectOrdinals`               | boolean | `true`           | Replace ordinals (1st) with superscript                         |
| `proofing.ignoreUppercase`                   | boolean | `true`           | Ignore words in UPPERCASE                                       |
| `proofing.ignoreWordsWithNumbers`            | boolean | `true`           | Ignore words that contain numbers                               |
| `proofing.ignoreInternetAddresses`           | boolean | `true`           | Ignore Internet and file addresses                              |
| `proofing.flagRepeatedWords`                 | boolean | `true`           | Flag repeated words                                             |
| `proofing.checkSpellingAsYouType`            | boolean | `false`          | Check spelling as you type                                      |
| `proofing.hideSpellingErrors`                | boolean | `false`          | Hide spelling errors                                            |
| `save.autoSave`                              | boolean | `true`           | Save changes automatically (AutoSave)                           |
| `save.autoRecoverIntervalMinutes`            | number  | `2`              | Save AutoRecover information every                              |
| `save.keepLastAutoRecoveredVersion`          | boolean | `true`           | Keep the last AutoRecovered version if I close without saving   |
| `save.defaultExportFormat`                   | string  | `'pptx'`         | Save files in this format                                       |
| `save.cacheRetentionDays`                    | number  | `14`             | Days to keep files in the local document cache                  |
| `save.clearCacheOnClose`                     | boolean | `false`          | Delete files from the local document cache when they are closed |
| `accessibility.subtitleLanguage`             | string  | `'auto'`         | (not shown in the dialog)                                       |
| `accessibility.showAccessibilityStatus`      | boolean | `true`           | Show accessibility status in the status bar                     |
| `accessibility.feedbackWithSound`            | boolean | `false`          | Provide feedback with sound                                     |
| `accessibility.soundScheme`                  | string  | `'modern'`       | Sound Scheme                                                    |
| `accessibility.showShortcutKeysInScreenTips` | boolean | `true`           | Show shortcut keys in ScreenTips                                |
| `accessibility.reducedMotion`                | boolean | `false`          | Reduced motion                                                  |
| `advanced.autoSelectEntireWord`              | boolean | `true`           | When selecting, automatically select entire word                |
| `advanced.allowTextDragAndDrop`              | boolean | `true`           | Allow text to be dragged and dropped                            |
| `advanced.maximumUndoSteps`                  | number  | `100`            | Maximum number of undos                                         |
| `advanced.useSmartCutAndPaste`               | boolean | `true`           | Use smart cut and paste                                         |
| `advanced.showPasteOptionsButton`            | boolean | `true`           | Show Paste Options button when content is pasted                |
| `advanced.imageDefaultResolution`            | string  | `'highFidelity'` | Default resolution                                              |
| `advanced.doNotCompressImages`               | boolean | `false`          | Do not compress images in file                                  |
| `advanced.chartPropertiesFollowDataPoint`    | boolean | `true`           | Properties follow chart data point                              |
| `advanced.recentPresentationsCount`          | number  | `50`             | Show this number of Recent Presentations                        |
| `advanced.showVerticalRuler`                 | boolean | `false`          | Show rulers                                                     |
| `advanced.showGrid`                          | boolean | `false`          | Show grid                                                       |
| `advanced.snapToGrid`                        | boolean | `false`          | Snap to grid                                                    |
| `advanced.disableHardwareAcceleration`       | boolean | `false`          | Disable hardware graphics acceleration                          |
| `advanced.disable3DRendering`                | boolean | `false`          | Disable 3D rendering (for performance)                          |
| `advanced.pixelateMosaicAnimation`           | boolean | `false`          | Show a mosaic effect for Pixelate transitions                   |
| `advanced.openDocumentsView`                 | string  | `'savedView'`    | Open all documents using this view                              |
| `advanced.slideShowShowMenuOnRightClick`     | boolean | `true`           | Show menu on right mouse click                                  |
| `advanced.slideShowShowPopupToolbar`         | boolean | `true`           | Show popup toolbar                                              |
| `advanced.slideShowPromptKeepInkAnnotations` | boolean | `true`           | Prompt to keep ink annotations when exiting                     |
| `advanced.slideShowEndWithBlackSlide`        | boolean | `true`           | End with black slide                                            |
| `advanced.printInBackground`                 | boolean | `true`           | Print in background                                             |
| `advanced.printHighQuality`                  | boolean | `false`          | High quality                                                    |
| `advanced.printUseMostRecentSettings`        | boolean | `true`           | Use the most recently used print settings                       |
| `advanced.printWhat`                         | string  | `'slides'`       | Print what                                                      |
| `advanced.printColorMode`                    | string  | `'color'`        | Color/grayscale                                                 |
| `advanced.printHiddenSlides`                 | boolean | `false`          | Print hidden slides                                             |
| `advanced.printScaleToFit`                   | boolean | `false`          | Scale to fit paper                                              |
| `advanced.printFrameSlides`                  | boolean | `false`          | Frame slides                                                    |
| `quickAccess.visible`                        | boolean | `true`           | Show Quick Access Toolbar                                       |
| `quickAccess.position`                       | string  | `'above'`        | Toolbar Position                                                |
| `quickAccess.showCommandLabels`              | boolean | `false`          | Always show command labels                                      |
| `trust.openInProtectedView`                  | boolean | `false`          | Open presentations in Protected View                            |
| `trust.allowExternalContent`                 | boolean | `true`           | Allow external content (remote images and media)                |
| `trust.confirmExternalHyperlinks`            | boolean | `true`           | Confirm before opening external hyperlinks                      |

### File tab pages (`backstage.hiddenPages`)

| Id        | Page    |
| --------- | ------- |
| `home`    | Home    |
| `new`     | New     |
| `open`    | Open    |
| `info`    | Info    |
| `save`    | Save    |
| `saveAs`  | Save As |
| `print`   | Print   |
| `share`   | Share   |
| `export`  | Export  |
| `close`   | Close   |
| `account` | Account |
| `options` | Options |

### File tab cards (`backstage.hiddenCards`)

| Id               | Card                            |
| ---------------- | ------------------------------- |
| `protect`        | Protect Presentation            |
| `inspect`        | Inspect Presentation            |
| `embedFonts`     | Embed Fonts                     |
| `signatures`     | Digital Signatures              |
| `versionHistory` | Version History                 |
| `saveAsPptx`     | PowerPoint Presentation         |
| `saveAsPpsx`     | PowerPoint Show                 |
| `saveAsPptm`     | Macro-Enabled Presentation      |
| `saveAsPpt`      | PowerPoint 97-2003 Presentation |
| `pdf`            | Create PDF                      |
| `png`            | Export current slide            |
| `video`          | Create a Video                  |
| `gif`            | Create an Animated GIF          |
| `json`           | Export as JSON                  |
| `copyImage`      | Copy as Image                   |
| `print`          | Print Presentation              |
| `share`          | Share with People               |

### Host context-menu commands

All five bindings accept `contextMenu.extraElementCommands` and `contextMenu.extraCanvasCommands`. These arrays contain `{ id, label, group?, disabled?, onSelect }` commands. Labels are plain text supplied by the host. `group` is `"top"` (the default) or `"bottom"`, separated from the built-in commands. IDs are isolated from built-in IDs; duplicate host IDs keep the first entry.

Element callbacks receive `{ slideIndex, elementIds }`; canvas callbacks receive `{ slideIndex }`. `slideIndex` is zero-based. Context is an immutable snapshot of the effective selection when the command was offered. `disabled` accepts a boolean or a predicate receiving that context. A disabled command cannot run. The menu closes before `onSelect` runs.

```ts
viewer.setCustomization({
	contextMenu: {
		extraElementCommands: [
			{
				id: 'send-to-chat',
				label: 'Send to chat',
				disabled: ({ elementIds }) => elementIds.length === 0,
				onSelect: ({ slideIndex, elementIds }) => sendToChat(slideIndex, elementIds),
			},
		],
		extraCanvasCommands: [
			{
				id: 'send-slide',
				label: 'Send slide',
				group: 'bottom',
				onSelect: ({ slideIndex }) => sendSlide(slideIndex),
			},
		],
	},
});
```

Use the customization prop/input/option or `setCustomization` / `updateCustomization` to replace commands live. Callbacks remain in memory and are not persisted in viewer settings. Disabling the corresponding menu also removes host commands.

### Element context menu (`contextMenu.hiddenElementCommands`)

| Id                       | Entry                |
| ------------------------ | -------------------- |
| `copy`                   | Copy                 |
| `cut`                    | Cut                  |
| `paste`                  | Paste                |
| `duplicate`              | Duplicate            |
| `edit-text`              | Edit Text            |
| `edit-points`            | Edit Points          |
| `bring-forward`          | Bring Forward        |
| `send-backward`          | Send Backward        |
| `bring-front`            | Bring to Front       |
| `send-back`              | Send to Back         |
| `ai-ask`                 | Ask AI about this    |
| `ai-fix`                 | Fix with AI          |
| `comment`                | Add Comment          |
| `hyperlink`              | Edit Hyperlink       |
| `table-insert-row-above` | Insert Row Above     |
| `table-insert-row-below` | Insert Row Below     |
| `table-delete-row`       | Delete Row           |
| `table-insert-col-left`  | Insert Column Left   |
| `table-insert-col-right` | Insert Column Right  |
| `table-delete-col`       | Delete Column        |
| `table-merge-selected`   | Merge Selected Cells |
| `table-merge-right`      | Merge Right          |
| `table-merge-down`       | Merge down           |
| `table-split`            | Split Cell           |
| `group`                  | Group                |
| `ungroup`                | Ungroup              |
| `crop`                   | Crop                 |
| `merge-union`            | Union Shapes         |
| `merge-combine`          | Combine Shapes       |
| `merge-fragment`         | Fragment Shapes      |
| `merge-intersect`        | Intersect Shapes     |
| `merge-subtract`         | Subtract Shapes      |
| `save-as-picture`        | Save as Picture...   |
| `edit-alt-text`          | Edit Alt Text...     |
| `size-and-position`      | Size and Position... |
| `format-shape`           | Format Shape...      |
| `format-chart-element`   | Format Chart Area... |
| `delete`                 | Delete               |

### Empty-canvas context menu (`contextMenu.hiddenCanvasCommands`)

| Id                  | Entry                |
| ------------------- | -------------------- |
| `paste`             | Paste                |
| `layout`            | Layout               |
| `reset-slide`       | Reset Slide          |
| `format-background` | Format Background... |
| `grid-and-guides`   | Grid and Guides      |
| `ruler`             | Ruler                |

### Edit Points menu (`contextMenu.hiddenEditPointsCommands`)

| Id                 | Entry            |
| ------------------ | ---------------- |
| `add-point`        | Add Point        |
| `delete-point`     | Delete Point     |
| `delete-segment`   | Delete Segment   |
| `open-path`        | Open Path        |
| `close-path`       | Close Path       |
| `smooth-point`     | Smooth Point     |
| `straight-point`   | Straight Point   |
| `corner-point`     | Corner Point     |
| `straight-segment` | Straight Segment |
| `curved-segment`   | Curved Segment   |
| `exit`             | Exit Edit Points |

### Editor shortcuts (`keyboard.disabled`, `keyboard.remap`)

| Id                   | Command                                                                |
| -------------------- | ---------------------------------------------------------------------- |
| `undo`               | Undo (Ctrl/Cmd+Z).                                                     |
| `redo`               | Redo (Ctrl/Cmd+Y, Ctrl/Cmd+Shift+Z).                                   |
| `copy`               | Copy the selection (Ctrl/Cmd+C).                                       |
| `cut`                | Cut the selection (Ctrl/Cmd+X).                                        |
| `paste`              | Paste (Ctrl/Cmd+V).                                                    |
| `duplicate`          | Duplicate the selection (Ctrl/Cmd+D).                                  |
| `delete`             | Delete the selection (Delete, Backspace).                              |
| `selectAll`          | Select every element on the slide (Ctrl/Cmd+A).                        |
| `group`              | Group the selection (Ctrl/Cmd+G).                                      |
| `ungroup`            | Ungroup (Ctrl/Cmd+Shift+G).                                            |
| `nudge`              | Move the selection with the arrow keys (disable only; not remappable). |
| `prevSlide`          | Previous slide (ArrowLeft with nothing selected).                      |
| `nextSlide`          | Next slide (ArrowRight with nothing selected).                         |
| `escape`             | Escape: leave the current mode (disable only; not remappable).         |
| `find`               | Find (Ctrl/Cmd+F).                                                     |
| `findReplace`        | Find and replace (Ctrl/Cmd+H).                                         |
| `toggleShortcuts`    | Keyboard-shortcut reference (?, Ctrl/Cmd+/).                           |
| `alignLeft`          | Align text left (Ctrl/Cmd+L).                                          |
| `alignCenter`        | Center text (Ctrl/Cmd+E).                                              |
| `alignRight`         | Align text right (Ctrl/Cmd+R).                                         |
| `alignJustify`       | Justify text (Ctrl/Cmd+J).                                             |
| `increaseFontSize`   | Increase font size (Ctrl/Cmd+Shift+>).                                 |
| `decreaseFontSize`   | Decrease font size (Ctrl/Cmd+Shift+<).                                 |
| `copyFormat`         | Copy formatting (Ctrl/Cmd+Shift+C).                                    |
| `pasteFormat`        | Paste formatting (Ctrl/Cmd+Shift+V).                                   |
| `newSlide`           | New slide (Ctrl/Cmd+M).                                                |
| `hyperlink`          | Insert or edit a hyperlink (Ctrl/Cmd+K).                               |
| `clearFormatting`    | Clear character formatting (Ctrl/Cmd+Space).                           |
| `cycleSelectionNext` | Select the next element (Tab).                                         |
| `cycleSelectionPrev` | Select the previous element (Shift+Tab).                               |
| `pasteSpecial`       | Paste Special (Ctrl/Cmd+Alt+V).                                        |

### Panels (`hiddenPanels`)

| Id                   | Region                                                                                          |
| -------------------- | ----------------------------------------------------------------------------------------------- |
| `statusBar`          | The status bar under the canvas (slide counter, zoom, view buttons).                            |
| `slidesPane`         | The slide thumbnail rail on the left.                                                           |
| `inspector`          | The format / properties inspector on the right.                                                 |
| `notes`              | The speaker-notes panel under the canvas.                                                       |
| `quickAccessToolbar` | The quick-access strip in the title bar (save, undo, redo, ...).                                |
| `titleBar`           | The title bar above the ribbon (file name, quick access, account).                              |
| `selectionPane`      | The Selection Pane and View > Selection Pane.                                                   |
| `findReplace`        | The Find and Replace panel, Home > Find / Replace and their shortcuts.                          |
| `accessibilityPane`  | The accessibility checker panel, Review > Check Accessibility and Help > Accessibility checker. |
| `comparePane`        | The Compare panel and Review > Compare.                                                         |

### Features (`disabledFeatures`)

| Id              | What it switches off                                                                         |
| --------------- | -------------------------------------------------------------------------------------------- |
| `ai`            | The AI assistant: toolbar toggle, chat panel, AI context-menu entries, AI Options page.      |
| `collaboration` | Real-time collaboration: Share and Broadcast buttons, dialogs and File pages.                |
| `comments`      | Commenting: the Add Comment context-menu entry and the Comments groups of Insert and Review. |
| `presentMode`   | Slide-show entry points: the Slide Show ribbon tab.                                          |
| `editPoints`    | Edit Points: the Edit Points context-menu entry and the point-editing mode it opens.         |

### Dialogs (`hiddenDialogs`)

| Id                   | What it removes                                                                   |
| -------------------- | --------------------------------------------------------------------------------- |
| `options`            | File > Options (the Settings dialog) and the File tab entry that opens it.        |
| `share`              | The Share dialog, its toolbar button, File > Share page and card.                 |
| `broadcast`          | The Broadcast dialog and its button.                                              |
| `print`              | The Print dialog, File > Print page and card.                                     |
| `export`             | File > Export page, its cards and the Export button.                              |
| `hyperlink`          | The Insert Hyperlink dialog: Insert > Link, the Link menu entry and its shortcut. |
| `equation`           | The Equation editor and Insert > Equation.                                        |
| `pasteSpecial`       | The Paste Special dialog and its shortcut.                                        |
| `slideTemplates`     | The slide template gallery and its entry at the end of the Home > New Slide menu. |
| `passwordProtection` | The password protection dialog and the File > Info Protect card.                  |
| `signatures`         | The digital signatures dialog and the File > Info Signatures card.                |

### Export formats (`hiddenExportFormats`)

| Id          | Format                                               |
| ----------- | ---------------------------------------------------- |
| `pdf`       | Export to PDF.                                       |
| `png`       | Export the current slide as PNG.                     |
| `video`     | Export the deck as a video.                          |
| `gif`       | Export the deck as an animated GIF.                  |
| `json`      | Export the parsed deck as JSON.                      |
| `copyImage` | Copy the current slide to the clipboard as an image. |

### Capabilities (`capabilities`)

| Id        | What `false` removes                                                                |
| --------- | ----------------------------------------------------------------------------------- |
| `edit`    | Editing. `false` opens the deck read-only whatever the edit prop says.              |
| `export`  | Exporting: the same as hiding the `export` dialog.                                  |
| `print`   | Printing: the same as hiding the `print` dialog.                                    |
| `share`   | Sharing: hides the `share` and `broadcast` dialogs.                                 |
| `copy`    | Copying content out: Copy, Cut, Save as Picture, their shortcuts and Copy as image. |
| `present` | Presenting: the same as disabling the `presentMode` feature.                        |
| `comment` | Commenting: the same as disabling the `comments` feature.                           |

### Replaceable regions (`regions`)

| Id                  | Region                                                                                                 |
| ------------------- | ------------------------------------------------------------------------------------------------------ |
| `inspector`         | The right-hand inspector column (the pane, not its open/close state).                                  |
| `slidesPane`        | The slide thumbnail rail on the left.                                                                  |
| `notes`             | The speaker-notes panel under the canvas.                                                              |
| `statusBar`         | The status bar under the canvas.                                                                       |
| `dialog.options`    | The body of File > Options (the Settings dialog). Context adds `options`, `setOption`, `resetOptions`. |
| `dialog.share`      | The body of the Share dialog.                                                                          |
| `dialog.broadcast`  | The body of the Broadcast dialog.                                                                      |
| `dialog.print`      | The body of the Print dialog.                                                                          |
| `backstage.home`    | The body of File > Home.                                                                               |
| `backstage.new`     | The body of File > New.                                                                                |
| `backstage.open`    | The body of File > Open.                                                                               |
| `backstage.info`    | The body of File > Info.                                                                               |
| `backstage.save`    | The body of File > Save.                                                                               |
| `backstage.saveAs`  | The body of File > Save As.                                                                            |
| `backstage.print`   | The body of File > Print.                                                                              |
| `backstage.share`   | The body of File > Share.                                                                              |
| `backstage.export`  | The body of File > Export.                                                                             |
| `backstage.close`   | The body of File > Close.                                                                              |
| `backstage.account` | The body of File > Account.                                                                            |
| `backstage.options` | The body of File > Options.                                                                            |

### Drawing tools (`hiddenDrawingTools`)

| Id              | Tool            |
| --------------- | --------------- |
| `freeformShape` | Freeform: Shape |
| `curve`         | Curve           |

<!-- customization-reference:end -->
