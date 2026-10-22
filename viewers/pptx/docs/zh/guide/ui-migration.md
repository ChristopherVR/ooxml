---
title: 共享 UI 迁移记录
description: 记录 PowerPoint 预览器迁移到共享 ooxml-ui/pptx 视图的证据：主题编辑器、功能区各选项卡、非功能区按钮、上下文菜单等，涵盖五个框架绑定。
---

# 共享 UI 迁移记录 {#shared-ui-migration-evidence}

迁移使用来自 `ooxml-ui/pptx` 的共享逻辑、open-shadow 视图和 CSS。原生适配器负责协调核心文档命令、历史记录、持久化和框架生命周期。不依赖 DOM 的编辑逻辑位于 `src/core/pptx/editor`。贡献者提出的 126 个工作单元是界面表面的清单，而不是要求公开 126 个自定义元素标签。

## 主题编辑器 (#364) {#theme-editor-364}

基线在实现之前已记录在对应 issue 中。所有截图使用同一份示例演示文稿、1440 x 900 的视口，以及“设计 > 编辑主题”入口。

| 绑定    | 基线位置                 | 迁移前                                                         | 共享编辑器                                                   |
| ------- | ------------------------ | -------------------------------------------------------------- | ------------------------------------------------------------ |
| React   | 功能区下方的右侧停靠面板 | [Before](/assets/ui-migration/theme-editor/react-before.png)   | [After](/assets/ui-migration/theme-editor/react-after.png)   |
| Vue     | 全窗口固定叠加层         | [Before](/assets/ui-migration/theme-editor/vue-before.png)     | [After](/assets/ui-migration/theme-editor/vue-after.png)     |
| Angular | 居中的库式模态框         | [Before](/assets/ui-migration/theme-editor/angular-before.png) | [After](/assets/ui-migration/theme-editor/angular-after.png) |
| Svelte  | 位于功能区内的编辑器     | [Before](/assets/ui-migration/theme-editor/svelte-before.png)  | [After](/assets/ui-migration/theme-editor/svelte-after.png)  |
| Vanilla | 功能区下方的右侧停靠面板 | [Before](/assets/ui-migration/theme-editor/vanilla-before.png) | [After](/assets/ui-migration/theme-editor/vanilla-after.png) |

共享编辑器停靠在编辑器主体一侧，宽度低于 768px 时变为底部面板。每个绑定都提供相同的十个预设、十二种颜色、字体字段、预览，以及“应用”“重置”和“关闭”。字段和预设只更新本地草稿；“应用”会向原生宿主发送一次完整的分离式编辑。“重置”恢复为宿主当前的主题。“关闭”和 Escape 会丢弃草稿并恢复打开者的焦点。次要的内联位置使用同一视图，但不会自动获取焦点。

属性、事件和生命周期约定记录在 `src/ui/src/pptx/web-components/README.md` 中。与框架无关的浏览器测试为 `e2e/pptx/theme-editor-migration.spec.ts`；它检查停靠、焦点、草稿丢弃、重置、应用、保存后的主题 XML、窄屏操作和强制颜色模式。适配器单元测试覆盖原生回调、禁用状态以及 React 18/19 兼容性。

## 功能区推进 (#363) {#ribbon-rollout-363}

计划中的每个功能区系列都已在全部五个绑定中交付。下面各节按交付顺序记录每一项。

| 系列                         | Issue | 状态                             |
| ---------------------------- | ----- | -------------------------------- |
| Home                         | #373  | 已交付，见 "Home"                |
| Insert                       | #374  | 已交付，见 "Insert"              |
| Transitions                  | #377  | 已交付，见 "Transitions"         |
| Animations                   | #378  | 已交付，见 "Animations"          |
| 独立的非功能区按钮和图标按钮 | #386  | 已交付，见 "非功能区按钮 (#386)" |

“幻灯片放映”“帮助”和“录制”这几批确立了命令/分组约定以及五个绑定的适配器；后续选项卡复用了它们。幻灯片内容的渲染仍然是单独的架构决策。本页是一份历史证据日志：记录的是每一批交付时测得的结果，而不是待办事项列表。

## 设计、审阅和上下文相关的库 {#design-review-and-contextual-galleries}

设计 (#376)、审阅 (#379)、形状格式 (#381)、图片格式 (#382)、表设计 (#383)、图表设计 (#384) 和 SmartArt 设计 (#385) 现在都在全部五个绑定中使用共享的分组、命令和库视图。主要的 UI 重构跟踪 issue (#342) 涵盖本页后文描述的控件原语和令牌。

“设计”命令的元数据控制可用性和打开状态。“审阅”使用七个规范分组和一个带键的分区视图，在原生状态变化时保持命令的焦点。库描述符提供预览和选择；共享视图负责弹出位置、键盘导航、关闭、触控目标和强制颜色样式。原生宿主仍然负责文档编辑、历史记录、选择、批注、对话框和保存。属性/事件约定记录在 `src/ui/src/pptx/web-components/README.md` 中。

可比较的截图使用相同的演示文稿和 1440 x 900 视口。基线文件位于 `/assets/ui-migration/ribbon-baseline/`；完成后的视图位于 `/assets/ui-migration/ribbon-after/`。文件名由绑定和选项卡组合而成，例如 `react-review.png` 和 `vue-shapeFormat.png`。

与框架无关的浏览器覆盖位于 `ribbon-contextual-migration.spec.ts`、`ribbon-gallery-migration.spec.ts` 和 `ribbon-review-migration.spec.ts`。它们实际执行图表和 SmartArt 的样式/颜色编辑并经过保存和重新加载，还覆盖审阅的拼写/语言/批注操作、自定义、焦点、窄屏布局、触控目标和强制颜色。保存/重新加载检查还发现并修复了缺失的 SmartArt 快速样式强度序列化、SmartArt 颜色定义更新以及图表库选择的重建问题。

## Draw {#draw}

Draw (#375) 使用一个受控的共享视图，涵盖选择、笔、荧光笔、橡皮擦、自由曲线、颜色和宽度。原生宿主负责指针手势、墨迹创建、擦除、历史记录和最近使用颜色的持久化。实时的自定义颜色输入会更新工具；确认的选择还会更新原生的最近使用颜色。五个绑定提供相同的宽度范围 (1 到 16) 和预设。

该视图在状态更新期间保留公开的自定义 id 和获得焦点的工具按钮。调色板会限制在视口内，在点击外部或按 Escape 时关闭，并在按 Escape 时恢复打开者的焦点。只读宿主会禁用绘制意图。主题令牌、强制颜色和粗指针目标遵循共享控件约定。可比较的截图是基线目录和完成后目录中的 `<binding>-draw.png`。

`e2e/pptx/ribbon-draw-migration.spec.ts` 覆盖键盘工具激活、原生墨迹创建、撤销/重做、InkML 保存/重新加载、自定义、令牌、触控目标和关闭。保存的 InkML 验证宽度和颜色；渲染出的路径或压力圆点验证重新加载后的笔画。现有的荧光笔和倾斜规格另行覆盖可变几何。

## View {#view}

View (#380) 使用一个受控的共享视图 `pptx-ui-ribbon-view`，用于“演示文稿视图”“母版视图”“显示”“缩放”和“窗口”分组。它负责分组标记、图标、标签、按下/选中状态和禁用状态。宿主提供查看器选项，并将类型化的 `view-request` 意图（command、option 或 guide）路由到原生处理程序。持久化的查看器选项（标尺、网格、参考线、对齐）、视图切换（普通、幻灯片浏览、大纲、阅读、幻灯片母版）、浏览器的取色器 (EyeDropper) 及其选择补丁、模板元素编辑以及编辑器历史记录，在每个绑定中仍保持原生。

现在五个绑定共享同一套控件：标尺、网格、参考线和“对齐网格”是复选框行；选择窗格、取色器、“对齐形状”、水平/垂直参考线和模板切换是带按下状态的命令。公开的自定义 id 保持不变，包括同时包裹两个参考线按钮的唯一 id `view.show.addGuide`。没有接入选择窗格或取色器的宿主会隐藏这些命令；隐藏 `zoom` 操作的宿主会从 DOM 中移除“缩放”分组。幻灯片母版、取色器和模板编辑在只读时被禁用。讲义母版、备注母版、缩放和宏仍然是禁用的占位项。

边界：全窗口的大纲、阅读和幻灯片浏览视图是原生叠加层，不属于本次变更。取色器的激活状态仅在宿主公开它的地方反映（React、Vue 和 Angular）。`e2e/pptx/ribbon-view-migration.spec.ts` 覆盖 id、键盘和指针切换、原生标尺效果、大纲/普通切换、模板状态、偏好持久化、自定义、触控目标、主题令牌和强制颜色。可比较的截图是基线目录和完成后目录中的 `<binding>-view.png`。

## Transitions {#transitions}

Transitions (#377) 使用一个受控的共享视图 `pptx-ui-ribbon-transitions`，用于“预览”“切换到此幻灯片”和“计时”分组：预设库、持续时间、声音（无、19 个内置声音、其他声音...）、“全部应用”、换片方式（单击鼠标时和之后）以及检查器开关。宿主通过共享的 `readRibbonTransitionDraft` 从当前幻灯片推导状态，并路由类型化的 `transitions-request` 意图。幻灯片变更、撤销历史、持久化、舞台上的切换预览重放、音频播放、读入保存流程的声音文件以及检查器面板仍保持原生。

一致性修复：现在每个绑定都显示相同的控件，当前预设带有按下状态，并在只读时禁止编辑（Angular 之前从不禁止）。持续时间实时提交，而“之后”的时间在失焦或按 Enter 时提交：React 和 Vue 过去每次按键都提交时间，Vanilla 和 Angular 则只在失焦时提交持续时间。公开的自定义 id 保持不变。效果选项的目录 id 在任何绑定中仍然没有对应控件。边界：Vanilla 和 Angular 宿主不跟踪检查器的打开状态，因此它们的检查器命令永远不会显示为按下。

`e2e/pptx/ribbon-transitions-migration.spec.ts` 覆盖 id、键盘激活、应用预设、持续时间、定时换片以及将内置声音应用到演示文稿、撤销/重做、保存和重新加载、跨幻灯片的“全部应用”、舞台预览、自定义、触控目标、主题令牌、焦点和强制颜色。`ribbon-control-effects.spec.ts`、`ribbon-compact-layout.spec.ts` 和 `effect-sound-gallery.spec.ts` 覆盖早期工作中的相同控件。可比较的截图是基线目录和完成后目录中的 `<binding>-transitions.png`。

## Animations {#animations}

Animations (#378) 使用一个受控的共享视图 `pptx-ui-ribbon-animations`，用于“预览”“动画”“动作路径”“高级动画”和“计时”分组。它负责分组标记、始终可见的进入/强调/退出和动作路径库、图标、标签、按下状态和禁用状态。宿主提供选择状态、编辑权限、动画窗格状态和翻译函数，并将类型化的 `animations-request` 意图（`add` 或 `command`）路由到原生处理程序。

原生宿主保留所有效果编辑（添加、删除和重新排序效果、触发器、计时、方向和重复）、Svelte 和 Vanilla 中支持拖动重新排序的动画播放顺序时间线、动画窗格/检查器的生命周期、画布内的预览播放、文档历史记录和持久化。现在五个绑定共享同一套控件：“退出效果”在所有地方都携带目录中的 `animations.advancedAnimation.addAnimation` id，检查器打开时动画窗格显示按下状态，Angular 的“效果选项”和“触发器”命令也像其他绑定一样打开动画面板。动画刷以及计时的开始/持续时间字段仍然是禁用的占位项；Vanilla 的面板命令仍然切换检查器。

公开的自定义 id 保持不变。边界：Svelte 和 Vanilla 的播放顺序时间线行以及检查器的动画面板是原生的，不属于本次变更。`e2e/pptx/ribbon-animations-migration.spec.ts` 覆盖 id、通过库添加的真实效果、文档级的动画窗格、撤销/重做、保存和重新加载、自定义、触控目标、焦点、主题令牌和强制颜色。可比较的截图是基线目录和完成后目录中的 `<binding>-animations.png`。

## Insert {#insert}

Insert (#374) 使用一个受控的共享视图 `pptx-ui-ribbon-insert`，用于“表格”“图像”“插图”（形状、自由曲线工具、图表、SmartArt）、“链接”（链接、动作）、“文本”（文本框、字段、页眉和页脚）、“符号”（公式）和“媒体”分组。它负责标记、图标、标签、选择器、按下状态的自由曲线以及禁用/只读限制；宿主提供查看器状态，并将类型化的 `insert-request` 意图（command、shapeType/shape、chartType/chart、freeform、actionButton、field）路由到原生处理程序。

在每个绑定中保持原生的内容：文档变更、撤销/历史记录和持久化，图像/媒体文件选择器，SmartArt 库，公式编辑器，超链接和页眉页脚对话框，自由曲线的预备状态和画布叠加层，以及（React、Vue、Angular 的）日期/时间格式选择器。公开的自定义 id 保持不变；页眉和页脚仍然没有 id。

迁移之前修复的一致性缺陷：Angular 的“插入”控件在只读演示文稿中从未被禁用，其“链接”也忽略选择状态；Svelte 和 Vanilla 对“动作”和“字段”使用原生 select 或弹出层，而 React、Vue 和 Angular 使用仅悬停触发的弹出层。现在五个绑定都共享点击和键盘菜单。“形状”按钮的字形在每个绑定中都跟随暂存的预设。对于“日期/时间”字段，Svelte 和 Vanilla 仍然直接插入当前日期（没有格式选择器），与之前相同。分组顺序现在遵循 PowerPoint（文本框移入“文本”分组），并且每个分组都有标题（新增的 `pptx.insert.group*` 键，包含英语、德语、西班牙语、法语和简体中文）。

`e2e/pptx/ribbon-insert-migration.spec.ts` 覆盖 id、通过键盘和指针插入文本框、形状和表格（含撤销/重做和保存/重新加载往返）、图表/动作/字段菜单、自由曲线的按下状态、链接的选择限制、原生的 SmartArt/公式对话框和图像文件选择器、自定义、触控目标、主题令牌和强制颜色。可比较的截图是基线目录和完成后目录中的 `<binding>-insert.png`。

## Home {#home}

Home (#373) 按分组系列交付，以便这个庞大的选项卡易于审阅。PR #359 已经把 Home 的图标图案 (`RIBBON_CONTROL_ICONS`)、目录 id 和编辑器外框布局 CSS 迁移到共享代码；但每个分组的标记、限制和回调仍然在各绑定中重复。下面每个系列现在都使用一个受控的共享视图 `pptx-ui-ribbon-home-<family>`，由它渲染按钮、图标、标签、按下状态和禁用状态，并发出一个携带公开控件 id 的类型化 `home-request` 意图（选择类操作另带 `value`）。宿主保留所有文档变更、历史条目和持久化。公开的自定义 id 保持不变。

- Clipboard：粘贴、剪切、复制和格式刷（整个 `home.clipboard` 分组）。现在五个绑定的限制逻辑完全一致；Angular 的格式刷之前在只读查看器中仍然可用，现在与其他四个一致。React 和 Vue 中“已复制/已剪切”的绿色闪烁只是视觉效果，没有保留。
- Font：`home.font` 中的字符条（加粗、倾斜、下划线、删除线、文字阴影、增大和减小字号、清除格式），加上字符间距、更改大小写，以及字体颜色和突出显示颜色的弹出层，还有作为 `font-picker` 的字体族和字号字段。各绑定如何计算切换和字号步进编辑保持不变（React 在点击时读取 run 级别的三态；各绑定之间的字号梯度不同）。
- Paragraph：`home.paragraph` 中的项目符号和编号（切换加共享的库）、减小和增大缩进、四种对齐方式、行距、文字方向和分栏。当查看器能读取到明确的对齐方式时，对齐会显示为按下状态。
- Editing：`home.editing` 中的查找、替换和选择菜单。
- Slides：整个 `home.slides` 分组（拆分式新建幻灯片、幻灯片模板、版式、重置、节），包括“新建幻灯片”和“版式”缩略图库。
- Drawing：形状和排列菜单、形状填充和形状轮廓的颜色弹出层、快速样式和形状效果。
- Arrange：对齐和分布、翻转、层次顺序、复制/删除共四个条，加上第二个格式刷、组合和取消组合、合并形状、裁剪（带纵横比菜单的拆分按钮）以及轮廓宽度微调框。

需要了解的行为变化：弹出层触发器带有 `aria-haspopup` 和 `aria-expanded`，并且全部在点击时打开（React 和 Vue 的悬停菜单以及 Angular 的悬停弹出层已不复存在）；菜单和颜色弹出层会在选择、按 Escape 或点击外部时关闭。现在每个文本附加项都在可编辑的文本选择上启用 (`canMut && canFormat`)，而不是仅在存在选择时启用。React 和 Svelte 中原本是英文字符串的菜单（行距、文字方向、分栏）现在可以翻译。字符间距提供共享的五个预设（Angular 较宽的列表已移除），“形状”在每个绑定中列出共享目录的前十二个预设（Vue 的自定义列表已移除），而“组合/取消组合”只存在于排列条中（不再在“绘图”的排列菜单中重复出现）。Svelte 的字号字段不再接受预设列表中不存在的输入字号（仍可显示这样的字号），与其他四个一致。Vue 中表格单元格现在也可以像 React 一样切换项目符号和编号。Angular 的排列以及 Svelte 的填充/轮廓/层次顺序现在遵循只读模式。

语言：Vue 和 Svelte 在响应式状态内，通过 `homeSnapshotTranslator` 为它们渲染的每个系列推导出共享的 `translate`，因此运行时切换语言会重新翻译这些条；React 和 Angular 在 `i18n`/`TranslateService` 变化时重新应用，Vanilla 则重建其外框。各绑定的单元测试和 `e2e/pptx/editor-controls-localization.spec.ts` 对此进行覆盖。

### 各控件的决策 {#decision-per-control}

此前原生的每个控件现在都渲染在共享元素中。适配器提供状态，并用各自可撤销的编辑函数处理类型化意图。

| 控件（自定义 id）                                  | 共享实现                                            | 适配器提供                                                                |
| -------------------------------------------------- | --------------------------------------------------- | ------------------------------------------------------------------------- |
| 字体族和字号 (`home.font.fontFamily`、`fontSize`)  | `pptx-ui-select` 上的 `font-picker` (`ribbon-font`) | 当前值、主题/嵌入/自定义字体；应用字体族和字号                            |
| 字符间距 (`characterSpacing`)                      | `font` 条，`pptx-ui-select` `ribbon-icon`           | 当前间距；应用 `characterSpacing`                                         |
| 更改大小写 (`changeCase`)                          | `font` 条，命令菜单                                 | 运行该绑定的文本变换（表格单元格使用 `textCaps` 提示）                    |
| 字体颜色、突出显示 (`fontColor`、`highlightColor`) | `font` 条，共享颜色弹出层                           | 主题颜色映射和最近使用项；应用颜色和主题引用，推送最近使用项              |
| 项目符号、编号 (`bullets`、`numbering`)            | `paragraph` 条：切换加内嵌库                        | 列表类型、库上下文；切换列表，应用库中的图块                              |
| 行距、文字方向、分栏                               | `paragraph` 条，`pptx-ui-select` `ribbon-icon`      | 当前值；应用样式补丁                                                      |
| 选择 (`home.editing.select`)                       | `editing` 条，命令菜单                              | “全选”处理程序                                                            |
| “新建幻灯片”下拉箭头、版式库                       | `slides` 条，版式图块库                             | 版式、当前版式、预览；通过 `layoutArtwork` 绘制缩略图图案；插入或应用版式 |
| 幻灯片模板 (`slideTemplates`)                      | `slides` 条按钮                                     | 打开原生对话框（对话框见 #342 和 #396）                                   |
| 形状、排列菜单 (`drawing`)                         | `drawing` 条，命令菜单                              | 插入形状，运行层次顺序命令                                                |
| 形状填充、形状轮廓                                 | `drawing` 条，共享颜色弹出层                        | 当前颜色和主题数据；应用 `shapeFillChange` / `shapeOutlineChange`         |
| 快速样式、形状效果                                 | `drawing` 条，内嵌 `pptx-ui-ribbon-gallery`         | 库上下文；应用图块                                                        |
| 第二个格式刷（排列中的 `formatPainter`）           | `arrange-painter` 条                                | 已启用/已选中状态；切换格式刷                                             |
| 组合、取消组合 (`group`、`ungroup`)                | `arrange-shape` 条                                  | 限制；运行该绑定的组合编辑                                                |
| 合并形状 (`mergeShapes`)                           | `arrange-shape` 命令菜单                            | 限制；运行合并操作                                                        |
| 裁剪 (`crop`)                                      | `arrange-shape` 拆分控件（主切换加菜单）            | 裁剪会话；切换它，或按比例裁剪、填充或适应                                |
| 轮廓宽度 (`outlineWidth`)                          | `arrange-shape` 数字输入                            | 当前宽度；应用 `strokeWidth`                                              |

刻意保持原生的部分及其技术原因：

- “幻灯片模板”对话框本身。它是一个模态对话框，其焦点管理归对话框迁移 (#342 和 #396) 负责；打开它的 Home 按钮是共享的。
- 每个版式缩略图内部绘制的图案。这属于幻灯片元素渲染（React 中的 `StaticElementRenderer`，Vue 和 Svelte 中的 `SlideStage`，Angular 中的 `pptx-element-renderer`，Vanilla 中的 DOM 渲染器），而幻灯片内容的渲染是单独的架构决策。缩略图外壳、几何、当前图块标记、键盘和弹出层是共享的，宿主通过 `layoutArtwork` 在已缩放的表面上绘制。
- 画布上下文菜单和缩略图栏中的版式库（React、Vue、Svelte 和 Angular 中的 `LayoutGalleryMenu`，以及 `ViewerMainContent`、`ViewerEditDialogs`）：上下文菜单留给 #393。
- Svelte 中编辑工具栏自带的文本格式分组（`FontExtrasGroup`、`FontFamilySelect`、`ShapeFormatGroup`）；它不属于 Home 功能区。

`e2e/pptx/ribbon-home-migration.spec.ts` 在全部五个绑定中覆盖 id、选择和剪贴板限制、带撤销和重做的真实复制/粘贴/剪切、格式刷、自定义、触控目标、主题令牌和强制颜色，现在还包括字体、颜色、段落、选择、形状填充/轮廓、轮廓宽度、排列菜单和版式库控件，并含保存和重新加载。本次变更的可比较截图位于 `/assets/ui-migration/home-remaining-before/` 和 `/assets/ui-migration/home-remaining-after/`（`<binding>-home.png`）；早期的 Home 截图仍在基线目录和完成后目录中。

## 非功能区按钮 (#386) {#non-ribbon-buttons-386}

下面的审计涵盖每一个非功能区按钮或图标按钮系列。它取自五个绑定的组件清单（React `viewer/components`、Vue `viewer/components`、Angular `viewer/*.component.ts`、Svelte `viewer/components`、Vanilla `viewer/ui`），并阅读了状态栏、只读横幅和粘贴选项的源码。除状态栏以外的系列是根据其文件结构和表中列出的源码分类的，并未迁移，因此“保持原生”是一个排期决策，而不是声称无法共享。功能区命令不在范围内：它们已经使用 `pptx-ui-ribbon-command` (#363)。只有当一个共享元素能够在不改变行为的前提下接管某个系列的标记、限制和回调时，该系列才会迁移。

| 系列                                             | 出现位置                                                                                                                             | 重复的标记或行为       | 决策                                                                                                     |
| ------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------ | ---------------------- | -------------------------------------------------------------------------------------------------------- |
| 状态栏（计数器、保存状态、备注、视图模式、缩放） | React `StatusBar.tsx`、Vue `StatusBar.vue`、Angular `status-bar.component.ts`、Svelte `StatusBar.svelte`、Vanilla `ui/status-bar.ts` | 是，五份手工构建的副本 | **迁移（本次变更）**：`pptx-ui-status-bar`。标签、限制和回调相同；宿主只映射状态和意图。                 |
| 只读横幅（仍然编辑、关闭、密码提示）             | 全部五个绑定中的 `ReadOnlyBanner`                                                                                                    | 是                     | 暂时保持原生。它包含带焦点、错误和忙碌状态的密码表单；将在后续作为独立元素迁移。                         |
| 粘贴选项工具栏                                   | 全部五个绑定中的 `PasteOptionsToolbar`                                                                                               | 是                     | 暂时保持原生。它锚定到画布元素并在捕获阶段关闭；需要先有共享的弹出层锚定约定。                           |
| 对话框页脚（取消、确定、应用、关闭）             | 每个绑定约 25 个对话框，位于不同的模态外壳中（React `useModalFocus`、Vue `ModalDialog.vue` 等）                                      | 是，但按对话框而异     | 保持原生。页脚与各绑定的焦点陷阱和关闭逻辑密不可分；必须先有共享的对话框外壳约定。                       |
| 上下文菜单：元素、画布、幻灯片缩略图、演示文稿   | 全部五个绑定                                                                                                                         | 是                     | 已迁移 (#393)：共享的 `pptx-ui-context-menu`；见下文“上下文菜单 (#393)”。浏览器和节菜单保持原生 (#397)。 |
| 演示工具栏和演讲者控制台工具栏                   | 全部五个绑定                                                                                                                         | 部分                   | 保持原生。放映时的叠加层，带自动隐藏、触控和全屏耦合。                                                   |
| 移动端底部栏和移动端顶部工具栏                   | React、Vue、Angular、Vanilla；Svelte 没有单独的底部栏                                                                                | 否，各绑定不同         | 保持原生：各绑定之间不一致。                                                                             |
| 文件后台导航和卡片                               | React、Angular、Svelte；Vue 和 Vanilla 没有对应项                                                                                    | 否，各绑定不同         | 保持原生：并非五个绑定都有。                                                                             |
| 幻灯片栏（缩略图、拖动重排、节行）               | 全部五个绑定                                                                                                                         | 部分                   | 保持原生。它承载幻灯片内容渲染和拖动状态，而 issue 将其排除在外。                                        |
| 标题栏和快速访问按钮                             | React `TitleBar.tsx`、Vue `TitleBar.vue`、Angular `title-bar.component.ts`、Svelte `TitleBar.svelte`、Vanilla `ui/title-bar.ts`      | 是                     | **迁移 (#394)**：`pptx-ui-title-bar`。统一的顺序、工具提示规则和限制表；由宿主负责的部分通过插槽提供。   |
| 检查器面板操作                                   | React 约 119 个文件、Vue 94、Svelte 74、Vanilla 63，Angular 更少但组件更大                                                           | 按面板而异             | 保持原生。按面板分别负责；通用约定是适用处的功能区命令、共享复选框和选择框。                             |
| 兼容性提示和协作状态指示器                       | 全部五个绑定（指示器通过插槽放入状态栏）                                                                                             | 是                     | 保持原生；提示的堆叠和中继重试由宿主负责。后续跟进。                                                     |
| 备注工具栏和备注面板按钮                         | React、Vue、Angular；Svelte 和 Vanilla 为内联                                                                                        | 是，行为有分歧         | **已在 #395 迁移**：`pptx-ui-notes-toolbar`；见下文“备注工具栏”一节。折叠标题和编辑器仍保持原生。        |

### 状态栏（第一批） {#status-bar-first-batch}

`pptx-ui-status-bar` 是底部一行的受控共享视图：“第 n 张，共 m 张”计数器、语言标签、保存指示器、备注开关、普通、幻灯片浏览和幻灯片放映按钮，以及缩放组（缩小、适应窗口的百分比读数、放大）。宿主提供 `StatusBarViewState`（通过宿主的翻译函数翻译），并将每次激活对应的一个类型化 `status-request` 意图路由到其原生处理程序。备注展开、视图切换、演示、幻灯片浏览器和缩放在每个绑定中仍保持原生。协作指示器通过名为 `collaboration` 的具名插槽放入。`resolveStatusBarSave` 是唯一的保存指示器规则（先看自动保存状态，再看脏标记），`statusBarViewMode` 将查看器模式映射到被按下的按钮。属性和事件约定记录在 `src/ui/src/pptx/web-components/README.md` 中。

为什么它不是功能区命令：状态栏是带实时读数的一行，包含一个插槽指示器，以及可见性取决于三个宿主操作（`hiddenActions` 中的 `zoom`、`notes` 和 `fullscreen`）的分组。逐个按钮做功能区命令会使限制逻辑和布局重复五遍。

行为变化，使五个绑定现在保持一致：

- 备注、普通、幻灯片浏览和幻灯片放映按钮在每个绑定中都暴露 `aria-pressed`。React、Vue 和 Angular 之前没有按下状态，Vanilla 仅在“备注”上有。
- 百分比读数在全部五个绑定中对辅助技术的名称都是 "Zoom to fit"。React、Vue 和 Angular 之前只暴露了工具提示，因此可见的 `100%` 就是其可访问名称。
- 保存文本使用同一条共享规则，包括 "Saved just now / n minutes ago"。Svelte 过去从不显示已保存文本，Vanilla 仍然显示其推送的标签。
- React 为“普通”和“幻灯片放映”读取查看器模式；Svelte 和 Angular 还会在浏览叠加层打开时按下“幻灯片浏览”。Vanilla 不跟踪浏览叠加层（保持不变）。
- 按钮至少为 24px 见方，在粗指针下增大到 44px；该行保持共享的 29px 最小高度。强制颜色使用系统颜色。
- React 的缩小和放大按钮过去各自可选；现在只要缩放组显示它们就会显示，缺少回调时为空操作。查看器总是提供全部三个回调。
- 宽度不超过 767px 时，该元素隐藏语言和保存文本。React、Vue 和 Angular 在手机上完全省略整个状态栏，Svelte 和 Vanilla 用 CSS 隐藏宿主；这些仍由宿主负责且保持不变。

`e2e/pptx/status-bar-migration.spec.ts` 在全部五个绑定中覆盖计数器、名称、按下状态、舞台上的缩放效果、键盘激活、幻灯片浏览器、面板自定义、触控目标、主题令牌和强制颜色。`chrome-shell-parity.spec.ts` 现在通过其 shadow root 测量该状态栏。适配器单元测试在每个绑定中覆盖状态映射、意图路由、限制、插槽投射和回调替换。

尚未迁移、并将在 #386 关闭之前拆分为有负责人的后续 issue 的内容：只读横幅、粘贴选项工具栏、使用共享对话框外壳的对话框页脚、演示工具栏和兼容性提示。这就是该 issue 仍保持打开的原因。

## 控件原语和令牌 (#342) {#control-primitives-and-tokens-342}

最初的观察结果（搜索框边框重复、文件导航间距不均、原生 select 弹出层、复选框强调色不一致）已在 #346 至 #349 中按表面逐个修复。这一轮把这个结果变为结构性的：每种控件类型一个共享原语，全部读取同一套设计令牌，并带有跨绑定检查，任何表面出现偏差都会失败。

令牌位于 `src/ui/src/pptx/web-components/control-tokens.ts`，通过 `tok(name)` 读取，它会生成 `var(--pptx-..., <default>)`。默认值在使用处解析，因此查看器根元素上的主题仍可影响它们。任何令牌都可以被宿主或主题覆盖。

| 分组   | 令牌                                                                                                                                      |
| ------ | ----------------------------------------------------------------------------------------------------------------------------------------- |
| 字段   | `--pptx-field-border`, `-border-focus`, `-bg`, `-fg`, `-placeholder`, `-radius`, `-height` (28px), `-height-lg` (40px), `-padding-x`      |
| 焦点环 | `--pptx-focus-ring-color`, `-width` (2px), `-offset` (2px)                                                                                |
| 密度   | `--pptx-space-1..4` (4, 8, 12, 16px), `--pptx-row-height` (28px), `--pptx-row-height-nav` (40px), `--pptx-touch-target` (44px)            |
| 复选框 | `--pptx-checkbox-size` (16px), `-size-touch` (22px), `-radius`, `-border`, `-bg`, `-accent`（主题主色）, `-accent-fg`                     |
| 单选框 | `--pptx-radio-dot-size` (8px), `-dot-size-touch` (10px)；边框、背景和强调色来自复选框令牌                                                 |
| 开关   | `--pptx-switch-width` (32px), `-height` (16px), `-track`, `-track-on`（主题主色）, `-thumb`, `-knob-size`, `-knob-offset`, `-knob-travel` |

| 类型   | 原语               | 仍保持原生之处                                                                                                                                                                                                  |
| ------ | ------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 搜索   | `pptx-ui-search`   | 无。标题栏 (28px) 和文件“最近使用的文件” (40px) 是同一个字段的两种变体。                                                                                                                                        |
| 选择框 | `pptx-ui-select`   | 无。选项、属性检查器、所有对话框（打印、自定义放映、超链接、文档属性、设置放映方式、日期/时间字段、页眉和页脚）以及功能区（动画时间线、插入选择器、绘制宽度、切换声音）在全部五个绑定中都使用应用自有的列表框。 |
| 复选框 | `pptx-ui-checkbox` | 无。每个对话框和面板中的复选框都是该原语。                                                                                                                                                                      |
| 单选框 | `pptx-ui-radio`    | 无。五个绑定中每一个原生 radio 输入（设置放映方式、放映幻灯片、打印设置、选择性粘贴、选项语言）都是该原语。Angular 选项窗格中的主题库和语言行是带有各自预览图案的 `role="radio"` 按钮，不是原生输入。           |
| 开关   | `pptx-ui-switch`   | 无。标题栏的自动保存开关就是该原语；它覆盖开关令牌，以保持标题栏尺寸所规定的 28 x 14px 轨道。五个绑定中不存在其他类似开关的切换项。                                                                             |

`pptx-ui-select` 是组合框/列表框，弹出层相对触发器定位，支持方向键、Home、End、PageUp、PageDown（每页八个选项）、预输入查找、Enter、Space、Escape 和 Tab 行为，支持 `aria-activedescendant`，在粗指针下触发器为 44px，并带有强制颜色样式。弹出层打开时按 Escape 只关闭弹出层：模态对话框（Svelte、Angular）过去会随之关闭，现在 `activateModalFocus` 避免了这一点。

新检查发现并修复的偏差：Vanilla 的标题栏搜索框是 24px 且背景不同（其他每个绑定都是 28px），并且它与 Vanilla 的文件搜索一起不显示占位符，因为该原语没有 `placeholder` 属性；Angular 选项的复选框，以及 Angular、Svelte 和 Vanilla 选项窗格把复选框设为 15px 而不是 16px。

证据：`/assets/ui-migration/primitives-before/` 和 `/assets/ui-migration/primitives-after/`，命名为 `<binding>-titlebar`、`-file-menu`、`-options` 和 `-inspector`。单元检查是 `control-primitives.test.ts`（令牌覆盖、搜索、复选框和选择框状态）和 `modal-focus.test.ts`；浏览器约定是 `e2e/pptx/ui-primitives-consistency.spec.ts`，在全部五个绑定中运行，与现有的 `search-field-focus`、`backstage-nav-spacing` 和 `web-controls` 规格并行。

### 对话框和面板控件 {#dialog-and-panel-controls}

五个绑定中每一个原生 `<select>` 和 `<input type="checkbox">` 现在都是 `pptx-ui-select` 或 `pptx-ui-checkbox`，通过检查器已在使用的同一套薄适配器（React `WebSelect`/`WebCheckbox`，Vue/Angular/Svelte 的自定义元素标签，Vanilla `createInspectorSelect`/`createInspectorCheckbox`）。值、`change` 事件、id、标签、可访问名称、`data-testid` 钩子和键盘行为保持不变；原先没有可访问名称的选择框现在有了名称。

| 表面                                 | 已转换                                                                 |
| ------------------------------------ | ---------------------------------------------------------------------- |
| 打印（对话框和设置面板）             | 打印内容、每页幻灯片数、范围、方向、颜色选择框；幻灯片加框和页眉复选框 |
| 自定义放映（对话框、功能区选择器）   | 活动放映选择框、逐张幻灯片复选框                                       |
| 超链接                               | 操作选择框                                                             |
| 文档属性（自定义选项卡）             | 属性类型和是/否值选择框                                                |
| 设置放映方式 / 放映幻灯片 / 放映选项 | 自定义放映选择框；循环、旁白、动画和字幕复选框                         |
| 日期/时间字段对话框                  | 格式选择框                                                             |
| 页眉和页脚、字体嵌入、查找           | 所有复选框                                                             |
| 选项功能区窗格、图表快速操作         | 所有复选框                                                             |
| 属性检查器面板                       | 其余的图表、效果、媒体、表格、文本、切换、主题覆盖和背景复选框         |
| 动画时间线 (Svelte、Vanilla)         | 触发器、形状、书签、方向、序列、曲线和重复选择框                       |
| 共享功能区视图                       | 插入形状/图表选择器、绘制宽度预设、切换的声音和换片复选框、动画的开始  |

保持原生的内容：无。字体名称和字号选择器已经使用带有已安装字体预览的共享选择框，并且没有任何文件类型选择器是 `<select>`。

选择框还新增了 PageUp 和 PageDown（每页八个选项，限制在最近的可用选项上），并且当宿主重新同步相同的值时，不再重建已打开的弹出层：React 在每次焦点变化时都会重新渲染“插入”功能区，这会替换指针下的选项并使点击丢失。

防护措施是 `src/ui/src/pptx/web-components/native-controls.test.ts`，它扫描每个绑定的源码（以及共享的 Web 组件）中是否存在原生 `<select>` 或 `<input type="checkbox">`，除非文件连同技术原因被列出，否则测试失败（该列表为空）。`e2e/pptx/dialog-controls.spec.ts` 在每个绑定中打开打印、超链接、设置放映方式、自定义放映和文档属性，如果页面显示任一原生控件则失败；`e2e/pptx/ui-primitives-consistency.spec.ts` 覆盖 PageUp/PageDown 行为。

每个绑定的前后对比截图位于 `/assets/ui-migration/dialog-controls-before/` 和 `/assets/ui-migration/dialog-controls-after/`，命名为 `<binding>-<dialog>`。之前的那组截图拍摄于 `bf656f3f2`，缺少当时某个绑定无法通过相同步骤打开的对话框。

仍在 #386 下保持打开、不在此处理的内容：对话框页脚、菜单、提示、横幅和工具栏。

### 单选按钮和自动保存开关 {#radio-buttons-and-the-autosave-switch}

`pptx-ui-radio` 取代每一个原生 `<input type="radio">`：设置放映方式（放映类型、放映幻灯片、换片）、放映幻灯片、打印设置（打印内容、幻灯片范围、方向、颜色模式；卡片标签内视觉上隐藏的 radio）、选择性粘贴以及 Vanilla 选项的语言列表。每个绑定中的适配器都很薄（React `WebRadio`，Vue、Angular 和 Svelte 的自定义元素标签，Vanilla `createInspectorRadio` 和 `appendRadioControl`）。名称、值和 change 行为保持不变；选择性粘贴中 Vue 的 `v-model` 改为 `:checked` 加 `@change` 的组合，而 Vanilla 的“所有幻灯片”和换片行过去只是通过给复选框设置 `type` 来假装是 radio，现在是真正的 radio 组。

带方向键、Home 和 End 的漫游 tabindex，以及每组仅一个 Tab 停靠点，已在 `control-primitives.test.ts` 中进行单元测试，并由 `e2e/pptx/dialog-controls.spec.ts`（设置放映方式）在全部五个绑定中检查。`ui-primitives-consistency` 断言 16px 的尺寸、圆形和焦点环，包括强制颜色模式。

`pptx-ui-switch` 是 `pptx-ui-title-bar` 中的自动保存开关。它保留可访问名称 "Toggle AutoSave"、`aria-checked`、Space 和 Enter、工具提示，以及宿主禁止自动保存时的惰性禁用状态。标题栏仍保持受控：触发 `change` 之后，它会恢复宿主的状态，因此拒绝操作不会让旋钮停在翻转后的位置。标题栏自身的 CSS 不再绘制开关；它根据共享的标题栏尺寸设置 `--pptx-switch-width`、`-height`、`-knob-*` 和 `-track`，因此轨道仍为 28 x 14px、旋钮为 10px，而粗指针下的 44px 点击区域位于该原语中。`chrome-shell-parity` 从开关的 shadow root 读取旋钮。`title-bar-migration` 新增了 Space/Enter 规格。

前后对比截图：`/assets/ui-migration/radio-switch-before/`（拍摄于 `0e4bacffd`，原生 radio 和标题栏按钮）和 `/assets/ui-migration/radio-switch-after/`（`<binding>-radios` 和 `<binding>-titlebar`）。

### #342 交付了什么 {#what-342-delivered}

每种控件类型一个原语：`pptx-ui-search`、`pptx-ui-select`、`pptx-ui-checkbox`、`pptx-ui-radio` 和 `pptx-ui-switch`，全部读取 `control-tokens.ts` 中的令牌来确定尺寸、强调色、焦点环和触控目标。五个绑定中每一个原生 select、checkbox 和 radio 都已转换，并且 `native-controls.test.ts` 对任何新增的原生控件都会失败（允许列表为空）。浏览器约定（`dialog-controls`、`ui-primitives-consistency`、`title-bar-migration`、`chrome-shell-parity`）在每个绑定中运行。不作声明的部分：Angular 选项的主题库和语言库仍是带预览图案的 `role="radio"` 按钮。

## 备注工具栏 (#395) {#notes-toolbar-395}

`pptx-ui-notes-toolbar` 是演讲者备注格式行的受控共享视图：加粗、倾斜、下划线、删除线、项目符号列表、编号列表、增大缩进、减小缩进、插入链接、打印备注，以及富文本/纯文本编辑器切换。它负责按钮、按钮顺序、图标、标签、启用状态、带漫游焦点的工具栏语义、触控目标、强制颜色和超链接弹出层。宿主保留 contenteditable 编辑器、段落和行内编辑命令、历史记录和持久化、打印以及折叠标题，并为每次激活路由一个类型化的 `notes-request` 意图。属性和事件约定记录在 `src/ui/src/pptx/web-components/README.md` 中。

### 变更前记录的差异 {#recorded-differences-before-the-change}

| 方面           | React                                                       | Vue                                          | Angular                        | Svelte                                       | Vanilla                                           |
| -------------- | ----------------------------------------------------------- | -------------------------------------------- | ------------------------------ | -------------------------------------------- | ------------------------------------------------- |
| 标签键         | `pptx.notes.*`                                              | `pptx.notesToolbar.*`                        | `pptx.notes.*`                 | `pptx.notes.*` 与 `pptx.notesToolbar.*` 混用 | `pptx.notes.*`                                    |
| 顺序           | 缩进、减少缩进                                              | 缩进、减少缩进                               | 缩进、减少缩进                 | 减少缩进、缩进                               | 缩进、减少缩进                                    |
| 纯文本模式     | 格式按钮保留但不起作用                                      | 相同                                         | 相同                           | 所有格式按钮隐藏                             | 格式按钮保留，并作用于隐藏的编辑器                |
| 模式切换标签   | "Plain editor" / "Rich editor"                              | "Plain" / "Rich"                             | "Plain editor" / "Rich editor" | "Plain" / "Rich"                             | "Plain editor" / "Rich editor"，带 `aria-pressed` |
| 打印限制       | 仅当宿主传入所有幻灯片时                                    | 始终                                         | 始终                           | 始终                                         | 始终                                              |
| 打印路径       | `NotesPrintDialog`                                          | 隐藏 iframe                                  | 隐藏 iframe                    | 隐藏 iframe                                  | 隐藏 iframe                                       |
| 工具栏限制     | `canEdit`                                                   | 仅当宿主可编辑时才挂载面板                   | 与 Vue 相同                    | 除非存在 `onupdate` 处理程序，否则隐藏       | 除非 `editable`，否则隐藏                         |
| 链接 UI        | 工具栏内弹出层                                              | 工具栏内弹出层（自有键集）                   | 工具栏内弹出层（自有键集）     | 两次 `window.prompt` 调用                    | 两次 `window.prompt` 调用                         |
| 图标           | Lucide 组件                                                 | Lucide 组件                                  | 内联 SVG                       | Lucide 组件                                  | 文本字形                                          |
| 工具栏语义     | 普通 `div`                                                  | 普通 `div`                                   | `role="toolbar"`，无漫游焦点   | 无角色的 `aria-label`                        | 普通 `div`                                        |
| 触控目标       | 约 24px                                                     | 约 24px                                      | 约 24px                        | 24px                                         | 24px                                              |
| 默认界面       | 仅当幻灯片有备注时为富文本；每次切换幻灯片都会重置该选择    | 桌面端为富文本，移动端为纯文本；保留         | 与 Vue 相同                    | 与 Vue 相同                                  | 与 Vue 相同                                       |
| 折叠和调整大小 | 标题按钮、`panelHeight` 和向下滑动的移动端面板              | 带箭头的标题按钮，`MobileSheet`              | 带箭头的标题按钮，移动端面板   | 不带箭头的标题按钮                           | 带文本箭头的标题按钮                              |
| 编辑器集成     | React 状态加重新填充编辑器的 effect；段落编辑使用 hook 状态 | 共享的 `notes-editor` 辅助函数，DOM 重新填充 | 与 Vue 相同                    | 与 Vue 相同                                  | 相同，但段落编辑使用上一次失焦时的状态            |

### 规范集合和决策 {#canonical-set-and-decisions}

- 一套键集 `pptx.notes.*`（增大缩进、减小缩进、项目符号列表、纯文本编辑器、富文本编辑器、插入链接、打印备注），加上现有的 `pptx.notesToolbar.ariaLabel`（工具栏名称）和 `pptx.common.cancel`。仅 Vue 使用的 `pptx.notesToolbar.*` 按钮键不再被读取。
- 一种顺序：字符格式、列表、先增大缩进再减小缩进、链接、打印，最后在行尾放编辑器切换。Svelte 之前遵循另一种顺序，现在与其余四个一致。
- 格式跟随编辑器：在纯文本编辑器中，格式按钮和链接按钮被禁用而不是隐藏（Svelte 过去会隐藏），也不再悄悄地失效。打印和编辑器切换保持启用。没有幻灯片时，所有按钮都被禁用。宿主无法编辑时，该行隐藏。
- 只要存在幻灯片就显示打印。React 的“打印”操作保留其全部幻灯片的对话框；其余四个仍通过隐藏 iframe 打印当前幻灯片，因此意图相同，但输出不同。
- 编辑器切换按钮命名为它将切换到的编辑器（富文本处于活动状态时显示 "Plain editor"）。它是不带 `aria-pressed` 的普通按钮，因此 Vanilla 的按下状态已移除。
- 链接弹出层移入该元素，因此 Svelte 和 Vanilla 不再使用 `window.prompt`。它是一个非模态对话框，会把自己放在该行上方（没有空间时放在下方），聚焦 URL，保留选中的文字作为显示文本，拒绝空 URL，在 Escape、取消或点击外部时关闭，并把焦点返回到“插入链接”按钮。提交时会恢复编辑器的选区，并发出一个带有规范化 URL 的 `link` 意图。
- 按钮在 `mousedown` 时阻止默认行为，使编辑器选区在指针按下后得以保留。宿主在运行命令之前先聚焦编辑器，因此键盘路径也能工作。
- React 的默认界面与其他绑定一样在桌面端为富文本，且选择不再因切换幻灯片而被重置。Vanilla 的列表和缩进命令先读取实时编辑器，因此自上次失焦以来输入的文本不再丢失。
- 富文本格式现在在每个绑定中都会持久化。Vue 和 Angular 过去只提交纯文本，因此加粗、列表和缩进会在下一次切换幻灯片和保存时丢失；现在它们在一个历史条目中提交文本及其片段（Angular 在 `update` 旁增加了 `notesCommit` 输出），在失焦时刷新，使切换幻灯片不会带走编辑内容，并且纯文本编辑或清空备注会丢弃过期的片段。当你返回到备注与上次保存的文本相同的幻灯片时，React 不再保留上一张幻灯片的草稿。
- 图标统一为一套内联 SVG。该行是只有一个 Tab 停靠点的 `role="toolbar"`；左、右、Home 和 End 在启用的按钮之间移动，且这些按键不会传到幻灯片快捷键。目标为 28px，粗指针下为 44px，强制颜色使用系统颜色。

刻意不变的部分：折叠标题、`panelHeight` 和移动端面板 (React)，contenteditable 编辑器及其历史记录，打印对话框和 iframe 路径，以及各绑定的移动端位置。每个绑定都保留其公开的 props、回调和测试类（`.pptx-vue-notes-toolbar`、`.pptxv-notes-toolbar`、`.pptx-svelte-notes-toolbar`、`#slide-notes-content`）。

| 绑定    | 迁移前                                                          | 迁移后                                                        |
| ------- | --------------------------------------------------------------- | ------------------------------------------------------------- |
| React   | [Before](/assets/ui-migration/notes-toolbar/react-before.png)   | [After](/assets/ui-migration/notes-toolbar/react-after.png)   |
| Vue     | [Before](/assets/ui-migration/notes-toolbar/vue-before.png)     | [After](/assets/ui-migration/notes-toolbar/vue-after.png)     |
| Angular | [Before](/assets/ui-migration/notes-toolbar/angular-before.png) | [After](/assets/ui-migration/notes-toolbar/angular-after.png) |
| Svelte  | [Before](/assets/ui-migration/notes-toolbar/svelte-before.png)  | [After](/assets/ui-migration/notes-toolbar/svelte-after.png)  |
| Vanilla | [Before](/assets/ui-migration/notes-toolbar/vanilla-before.png) | [After](/assets/ui-migration/notes-toolbar/vanilla-after.png) |

`e2e/pptx/notes-toolbar-migration.spec.ts` 在全部五个绑定中覆盖按钮顺序和名称、能在切换幻灯片后保留的真实加粗、项目符号和缩进编辑、不使用 `window.prompt` 的链接弹出层、Escape 和空 URL 处理、漫游焦点、被禁用的纯文本模式、打印、主题令牌、强制颜色以及 44px 触控目标；`mobile-notes.spec.ts` 检查默认的纯文本界面及其触控目标。适配器单元测试在每个绑定中覆盖状态映射、意图路由、限制和回调替换。

## 幻灯片栏、节和浏览器操作 (#397) {#slide-rail-section-and-sorter-actions-397}

幻灯片栏本身（缩略图渲染、虚拟化、拖动状态）仍保持原生。所做的决定是幻灯片栏、节标题和浏览器图块各自暴露哪些操作，因为审计发现每个绑定背后都是不同的产品行为。变更前记录的证据：

- **幻灯片栏常驻行。** React、Svelte 和 Vanilla：一个“添加幻灯片”页脚。Vue：仅平铺式演示文稿有页脚（分节演示文稿没有），而 `SlidesPaneControls.vue`（添加、复制、删除）存在但从未被挂载。Angular：页脚加上每个缩略图上的悬停工具栏（复制、删除、上移、下移），且没有拖动重新排序。
- **节标题。** React：带内联重命名的弹出菜单 (`pptx.sections.*`)。Vue：悬停按钮 (`pptx.sectionList.*`)、每个分组一个“+ 添加节”按钮以及内联重命名。Angular、Svelte、Vanilla：悬停按钮以及 `window.prompt` 重命名。
- **浏览器图块菜单。** React：复制、粘贴（需有剪贴板）、复制幻灯片、隐藏和显示、删除，并带数量后缀。Vue、Angular、Svelte：复制幻灯片、隐藏/显示、删除 (`pptx.slideMenu.*`)。Vanilla：没有菜单，每张卡片上有内联的复制幻灯片、隐藏/显示和删除按钮。
- **分节幻灯片栏上的缩略图菜单。** React、Angular、Svelte 和 Vanilla 中存在；Vue 中缺失（其节列表没有右键菜单，也没有多选）。

### 决策 {#decision}

以 PowerPoint 为参照：其幻灯片窗格没有每个缩略图上的按钮，也没有按钮条，其节标题和浏览器都是右键菜单。因此：

- **幻灯片栏。** 在固定页脚中只有一个常驻操作 **Add Slide** (`pptx.sections.addSlide`)，平铺式和分节演示文稿皆如此。幻灯片上的其他一切操作都在缩略图右键菜单（已在 `slide-pane-context-menu.ts` 中共享）和键盘上（Enter 在其后插入，Delete）。重新排序在全部五个绑定中都是拖放（Angular 新增了该功能；其每个缩略图上的上移/下移按钮已移除）。Vue 的分节幻灯片栏新增页脚、缩略图菜单和 Ctrl/Shift 多选，并且未挂载的 `SlidesPaneControls.vue` 被删除。
- **节标题。** 右键或键盘上下文菜单键触发的 `role="menu"` 弹出层，包含重命名、删除、上移、下移和在其后添加节（即 React 的列表，与 PowerPoint 一致）。上移在第一个节上禁用，下移在最后一个节上禁用。重命名是标题中的内联文本字段（Enter 或失焦提交非空名称，Escape 取消），所有绑定皆如此；没有任何绑定打开 `window.prompt`。“在其后添加节”会让新节从该节最后一张幻灯片之后的那张幻灯片开始，并限制在最后一张幻灯片内。悬停按钮、Vue 每个分组的“+ 添加节”按钮和 `pptx.sectionList.*` 键不再被幻灯片栏使用；所有地方的标签键都是 `pptx.sections.*`。
- **浏览器图块。** 右键触发的 `role="menu"` 弹出层，包含复制、粘贴（仅在复制之后）、复制幻灯片、隐藏或显示、删除。隐藏和显示与幻灯片栏菜单一样是同一个切换项（仅当所有选中的幻灯片都已隐藏时才显示“显示”）；当删除会移除所有幻灯片时禁用删除；多选时会在复制、复制幻灯片、隐藏/显示和删除后追加 " (n)"。键为 `pptx.slideSorter.contextMenu.*`。Vanilla 的内联每卡按钮已移除。复制和粘贴也响应 Ctrl+C 和 Ctrl+V（共享的浏览器键映射）。

共享逻辑位于 `ooxml-ui/pptx`：`buildSectionContextMenuEntries` 和 `sectionAddAfterSlideIndex`（`section-context-menu.ts`），`buildSlideSorterContextMenuEntries`、`slideSorterContextMenuLabel`、`slideSorterPasteIndexes` 和 `SLIDE_RAIL_FOOTER_ACTIONS`（`slide-sorter-context-menu.ts`）。每个绑定以原生方式渲染这些条目（React 和 Vue 复用其菜单组件，Angular 有 `pptx-section-context-menu`，Svelte 有 `SectionContextMenu.svelte` 和 `SlideSorterContextMenu.svelte`，Vanilla 有 `section-context-menu.ts` 和 `slide-sorter-context-menu.ts`）；共享的上下文菜单元素落地后，这些菜单会并入其中。

浏览器的选择和剪贴板状态现在位于 `slide-sorter-state.ts` (#402)。全部五个适配器都支持 Ctrl/Cmd 切换、Shift 范围选择、全选、菜单中的选择计数、多幻灯片复制/粘贴和缩略图缩放。单击选中；双击打开幻灯片。Escape 会先收起多选，再关闭。粘贴按演示文稿顺序复制被复制的源幻灯片，与现有 React 行为一致。`e2e/pptx/slide-sorter-selection.spec.ts` 对每个绑定运行此约定。节颜色、折叠和拖动状态仍保持原生。

`e2e/pptx/slide-rail-menus-parity.spec.ts` 在全部五个绑定中驱动“添加幻灯片”页脚（平铺式和分节）、节菜单（命令、角色、无内联按钮）、无提示框的内联重命名、上移/下移的限制和重新排序，以及浏览器菜单（命令、先复制再粘贴、先隐藏再显示）。

## 检查器的重置和清除操作 (#398) {#inspector-reset-and-clear-actions-398}

检查器面板仍保持原生（每个面板通过各自的编辑器和历史记录 API 提交）。审计发现的是六个操作的一致性缺口，变更前按翻译键记录如下：Svelte 中“重置图片” (`image.resetImage`) 缺少标签和限制约定（写死的英文），而在其他地方有三种不同的限制方式；“重置裁剪”在 Svelte 和 Vanilla 中缺失；“重置修剪”在 Angular、Svelte 和 Vanilla 中缺失；“清除系列颜色”在 Svelte 和 Vanilla 中缺失；“清除背景”在 Svelte 和 Vanilla 检查器中缺失（二者仅在“设置背景格式”功能区停靠面板上才有）。

`ooxml-ui/pptx` 现在在 `inspector-reset-actions.ts` 中拥有这份约定：标签键、限制逻辑（`imageResetState`、`cropResetState`、`mediaTrimResetState`、`seriesColorClearState`、`slideBackgroundClearState`），以及每个操作所应用的确切补丁（`imageResetPatch`、`cropResetPatch`、`mediaTrimResetPatch`、`slideBackgroundClearPatch`）。规范行为（改变了部分绑定）：

- **重置图片** 对图片显示，且仅在宿主可编辑并且存在效果或裁剪为形状的覆盖时启用。它清除 `imageEffects` 和 `cropShape`，并保留裁剪（PowerPoint 保留裁剪；由“重置裁剪”负责）。Vue、Angular、Svelte 和 Vanilla 过去也会清除裁剪；Vue 在未修改前隐藏该按钮，React 则从不禁用它。
- **重置裁剪** 对图片显示，在可编辑且图片未被 `noCrop` 锁定时启用，并将四个内缩量归零。已加入 Svelte 和 Vanilla。
- **重置修剪** 仅对带有修剪的可编辑媒体显示。已加入 Angular、Svelte 和 Vanilla。
- **清除系列颜色** 仅在可编辑且该系列有自己的颜色时，按系列显示。已加入 Svelte 和 Vanilla（Vanilla 只有一个系列选择器，因此为所选系列提供一个按钮）。现在每个绑定都通过可访问名称暴露该操作（React、Vue 和 Angular 之前只有标题或字形）。在 Vanilla 中现在只有颜色控件会写入系列颜色；编辑趋势线或误差线过去会把色板颜色写到系列上。
- **清除背景** 在幻灯片有颜色、图片、渐变或图案时显示，可编辑时启用，并清除这四项（React、Vue 和 Angular 过去跳过图案）。Svelte 和 Vanilla 在无选择状态的检查器中新增带颜色和此操作的“背景”卡片，并且 Svelte 功能区停靠面板的清除按钮现在显示 "Clear Background" 而不是 "Default"。
- 每个操作都是一个撤销步骤。

标签使用现有的键，已翻译为德语、西班牙语、法语和简体中文；无需新增键。Svelte 的“替换”和“重置图片”标签现在使用 `pptx.image.replaceImage` 和 `pptx.image.resetImage`。

`e2e/pptx/inspector-reset-actions.spec.ts`（夹具 `inspector-reset-actions.pptx`）在全部五个绑定上检查这五个操作、它们的限制和撤销。

### 标题栏和快速访问 (#394) {#title-bar-and-quick-access-394}

`pptx-ui-title-bar` 是顶部外框行和可选的功能区下方条带的受控共享视图。宿主提供 `TitleBarViewState`（通过 `buildTitleBarState` 构建），并将类型化的 `toggle-autosave`、`save`、`undo`、`redo`、`quick-command` 和 `command-search` 事件路由到其原生处理程序。协作指示器和账户部分由宿主负责，通过具名的 `collaboration` 和 `account` 插槽放入。属性、事件和键盘约定记录在 `src/ui/src/pptx/web-components/README.md` 中。

五个绑定在此前的做法，按本次变更中的审计结果：

| 方面               | React                                          | Vue                          | Angular                                | Svelte                                                | Vanilla                                                           |
| ------------------ | ---------------------------------------------- | ---------------------------- | -------------------------------------- | ----------------------------------------------------- | ----------------------------------------------------------------- |
| 布局和令牌         | Tailwind `TITLE_BAR_CLASSES`                   | Tailwind `TITLE_BAR_CLASSES` | Tailwind `TITLE_BAR_CLASSES`           | 来自 `titleBarStyleAttr` 的作用域 CSS；名称最大 200px | 来自 `TITLE_BAR_METRICS` 的 CSS-in-TS；名称最大 180px，搜索 320px |
| 存在的按钮         | 仅在传入 `onSave` 时有保存；撤销、重做、附加项 | 与 React 相同                | 始终有保存；撤销、重做、附加项         | 始终有保存；手绘的保存、撤销、重做字形                | 按配置顺序列出每个已配置的 id，因此没有固定的三件套               |
| 附加项和工具提示   | 仅标签，忽略 `showCommandLabels`               | ScreenTip 解析器             | ScreenTip 解析器；拼写检查画成搜索字形 | 仅标签；未知图标回退为播放                            | ScreenTip 解析器；拼写检查没有字形；撤销没有操作文本              |
| 自动保存开关       | 始终启用                                       | 始终启用                     | 始终启用                               | 始终启用；“宿主禁用”的状态键无法触达                  | 宿主禁止自动保存时惰性（唯一如此的绑定）                          |
| 搜索字段限制       | `mode` 为编辑或母版                            | `mode` 为编辑或母版          | 可编辑                                 | 可编辑                                                | 可编辑                                                            |
| 搜索结果           | 共享过滤、上限 8、类别、“在幻灯片中查找”行     | 与 React 相同                | 自有组件，行相同                       | 与 React 相同，失焦宽限 120ms                         | 本地三项列表；没有类别、上限或回退行                              |
| 用户和在线状态区域 | 无                                             | 无                           | 无                                     | 无                                                    | AI 开关直接追加到该行                                             |
| 协作插槽           | 无（指示器位于状态栏）                         | 无                           | 无                                     | 无                                                    | 无                                                                |
| 窗口控件           | 无（查看器嵌入宿主页面；窗口归宿主所有）       | 无                           | 无                                     | 无                                                    | 无                                                                |
| 功能区下方         | `TitleBarQuickExtras`                          | `TitleBarQuickAccess`        | `QuickAccessStripComponent`            | `QuickAccessToolbar.svelte`                           | 迁移实时的条带节点                                                |
| 第二个分隔线       | 启用条带时                                     | 启用条带时                   | 始终                                   | 启用条带时                                            | 条带停靠在下方时隐藏                                              |
| 窄宽度             | 768px 以下隐藏；不换行也不处理溢出             | 相同                         | 相同                                   | 横屏手机上也隐藏                                      | 横屏手机上也隐藏                                                  |

现在五个绑定共享的决策：

- 一种顺序：保存、撤销、重做，然后按“文件 > 选项”的顺序列出已配置的附加项。
- 一条工具提示规则：`aria-label` 是标签，`title` 是 `screenTip(...)`，撤销和重做在已知时会指出待处理的操作。附加项遵循 `showCommandLabels`。
- 一张限制表：自动保存、条带、状态和搜索仅在编辑时显示（编辑或母版模式，且演示文稿可编辑）。`showSave`、`showUndo` 和 `showRedo` 可隐藏单个按钮。仅当条带有按钮时才出现第二个分隔线。
- 每个目录命令使用同一套图标（包括拼写检查），取自共享的受信路径。
- 一个搜索下拉框：共享过滤、最多八行、类别列，以及在宿主可以搜索内容时显示的“在幻灯片中查找”行。方向键在结果间移动。
- 条带是带有漫游焦点的具名工具栏。触控下按钮为 44px，强制颜色使用系统颜色。该行在每个绑定中都在 768px 以下隐藏。

发生变化的行为：

- React 和 Vue 过去在任何编辑模式视图中都显示搜索框；现在与其他三个一样，需要可编辑的演示文稿。
- 当宿主禁止自动保存时，自动保存开关在全部五个绑定中都是惰性的（禁用，并显示“宿主已禁用”的工具提示）；此前只有 Vanilla 如此。
- Vanilla 现在总是先显示保存、撤销和重做；省略它们的已配置列表不再隐藏它们，这与其他四个一致。它的搜索框现在是共享的那个，列出共享的命令目录，通过功能区自己的界面运行每个 id（SmartArt、公式、浏览主题和幻灯片大小会打开对应的功能区控件），并提供“在幻灯片中查找”行，它会切换“查找和替换”面板。
- AI 开关被放入 `account` 插槽，并且 Vanilla 的搜索框现在会显示其占位符（它之前赋值的属性字段从不读取，所以显示为空）。
- 现在在 React 中隐藏 `quickAccessToolbar` 面板也会移除保存、撤销和重做；此前只隐藏附加项。

现在五个绑定都会把待处理的撤销和重做操作传给工具提示（Vue、Svelte 和 Vanilla 新增了这条通路；目前没有任何绑定中的编辑操作记录标签，所以今天每个工具提示都只显示 "Undo" 和 "Redo"）。Svelte 还传递自动保存被禁用的原因，因此宿主设置 `autosave: false` 或缺少文件路径时，所有地方显示相同的状态文本。`review.language` 在任何绑定中都没有对应操作。

`e2e/pptx/title-bar-migration.spec.ts` 在全部五个绑定中覆盖名称、顺序、自动保存开关、漫游焦点和 Tab 顺序、搜索结果、面板自定义、窄宽度、主题令牌、触控目标和强制颜色。`chrome-shell-parity.spec.ts` 现在通过其 shadow root 测量该栏，并跨绑定比较命令搜索行。适配器单元测试在每个绑定中覆盖状态映射、事件路由、限制、插槽和功能区下方的位置。前后对比截图位于 `docs/public/assets/ui-migration/title-bar-before` 和 `title-bar-after`。

## 上下文菜单 (#393) {#context-menus-393}

`pptx-ui-context-menu` 是全部五个绑定中元素、空白画布、幻灯片缩略图和幻灯片放映菜单（以及 Vanilla 的只读 AI 菜单）背后的一个受控共享元素。每个绑定保留其条目列表（`buildContextMenuEntries`、`buildCanvasContextMenuEntries`、`buildSlidePaneContextMenuEntries`、`getPresentationContextMenuSections`）、宿主自定义、限制和命令分派（剪贴板、表格、裁剪、合并、检查器焦点）；适配器把这些条目映射为行，并路由类型化的 `menu-request { id }` 和 `menu-close` 事件。属性和事件约定记录在 `src/ui/src/pptx/web-components/README.md` 中。浏览器菜单和节菜单在各绑定中各不相同（见节和浏览器一致性 issue），仍保持原生；Vue 的浏览器复用通用的 `ContextMenu.vue`，因此同样获得共享的外观和键盘模型。

变更前从五个绑定读取的证据：

| 方面                         | React              | Vue                                                | Angular                | Svelte             | Vanilla                         |
| ---------------------------- | ------------------ | -------------------------------------------------- | ---------------------- | ------------------ | ------------------------------- |
| 元素 z-index                 | 120（遮罩 119）    | 120                                                | 9000                   | 120（遮罩 119）    | 1150                            |
| 幻灯片放映 z                 | 1300（遮罩 1299）  | 120，位于 2147483000 的全屏叠加层之下              | 9000                   | 96（遮罩 95）      | 65                              |
| 幻灯片放映限位               | 无                 | 与元素菜单共享                                     | 无                     | 无                 | 共享的双侧限位                  |
| 点击外部关闭                 | 吞掉点击的全屏遮罩 | window 上的捕获监听器                              | `document:pointerdown` | 吞掉点击的全屏遮罩 | document 上的捕获 `pointerdown` |
| 幻灯片窗格钩子               | 标记和名称         | 没有 `data-pptx-slide-pane-context-menu`，没有名称 | 标记和名称             | 标记和名称         | 标记和名称                      |
| 删除着色                     | 有                 | 元素菜单失去了它                                   | 有                     | 有                 | 缩略图菜单上没有                |
| 幻灯片放映名称               | 无                 | 无                                                 | 无                     | 无                 | 无                              |
| 分组标题                     | 指针和屏幕标题     | 无                                                 | 无                     | 无                 | 无                              |
| 方向键、预输入查找、漫游焦点 | 无                 | 无                                                 | 无                     | 无                 | 无                              |

由此暴露并修复的真实缺陷：

- Vue 的幻灯片放映菜单被传送到 `<body>` 且 z-index 为 120，而放映是全屏叠加层，因此菜单渲染在幻灯片后面，无法点击（前目录中的 `vue-presentation.png` 是一个空白的幻灯片片段）。共享菜单现在在全屏子树内就地渲染，位于所有叠加层之上。
- Vue 的幻灯片缩略图菜单没有可访问名称，也没有其他四个绑定输出的 `data-pptx-slide-pane-context-menu` 钩子；其元素菜单丢失了删除的危险色调。
- 幻灯片放映菜单在任何绑定中都没有可访问名称，并且只有 React 绘制了其“指针选项”和“屏幕”标题。现在全部五个绑定都共享二者，作为 `group`。
- 幻灯片放映中的 Escape 只关闭已打开的菜单；它在放映自身的按键处理之前被消费。

行为变化，使五个绑定现在保持一致：

- 行是 `menuitem` 或 `menuitemcheckbox`；分隔线是 `separator`；界面是位于该元素开放 shadow root 内的具名 `role="menu"`。宿主元素携带 `data-pptx-*` 钩子，因此宿主上的现有选择器仍然有效。
- 打开时聚焦第一个启用的行。方向键循环，Home 和 End 跳转，预输入查找匹配标签前缀、重复字母则循环，并跳过禁用的行。悬停某行会将漫游焦点移到该行。Enter 和 Space 激活。Escape、点击外部和 Tab 关闭；除非命令移动了焦点，否则焦点返回到打开者。
- 点击外部不再吞掉点击（React 和 Svelte 过去会）：菜单关闭，点击继续传递，Vue、Angular 和 Vanilla 此前已是如此。
- 每个菜单都被限制在窗口内，高于窗口时滚动。React、Angular 和 Svelte 的幻灯片放映菜单此前没有限位。
- 一份层叠约定：编辑器为 9000，幻灯片放映为 2147483001。
- 行高为 28px，粗指针下增至 44px。颜色来自共享主题令牌；强制颜色使用系统颜色且无阴影。
- Vanilla 的只读 AI 菜单现在使用同一元素，因此失去了它的两个图标。没有任何绑定有子菜单，所以模型中也没有；Vanilla 的画布和缩略图“版式”列表是第二个共享菜单，而不是嵌套菜单。

保留的原生代码：条目列表和宿主自定义、`context-menu-dispatch`、`editor-context-menu-dispatch`、`element-context-menu-commands` 以及各绑定中等价的分派、命中测试和选择，以及 React、Vue、Angular 和 Svelte 从“版式”打开的版式库。“编辑顶点”的节点菜单锚定在各绑定的编辑顶点叠加层内部，而不是指针处，不属于此清单。

可比较的截图使用 1440 x 900 的示例演示文稿，围绕菜单裁剪：`/assets/ui-migration/context-menu/before/<binding>-<menu>.png` 和 `/after/`，涵盖 `element`、`canvas`、`slide-rail` 和 `presentation`（例如 `react-element.png`）。

`e2e/pptx/context-menu-migration.spec.ts` 每个绑定运行一次，覆盖标记和角色、第一行焦点、方向键、Home、End 和预输入查找、Enter 激活、Escape 和点击外部关闭并能干净地重新打开、画布开关、缩略图菜单、位于叠加层之上并带分组的幻灯片放映菜单、右下角的限位、主题令牌、强制颜色和 44px 触控行。现有的 `context-menu-parity`、`canvas-context-menu-parity`、`slides-pane-parity`、`host-context-menu-commands` 和 `edit-points` 规格通过共享元素的 shadow root 读取菜单（`e2e/pptx/support/context-menu.ts`）。共享单元测试覆盖元素约定、行映射器以及各绑定的适配器。

## 非功能区系列 (#386) {#non-ribbon-families-386}

#386 的第一轮根据文件清单对十二个非功能区系列进行了分类，并只迁移了状态栏。这一轮在决定之前，阅读了全部五个绑定中每个系列的源码。当某个系列的标记、限制和回调在各绑定中相同，或者仅存在可以在不改变用户可执行操作的前提下规范化的偏差时（与状态栏使用的规则相同），该系列即被迁移。只有当文件证据表明各绑定实现的是不同功能时，系列才保持原生，并且其中每一项都有一个专门的后续 issue。功能区命令不在范围内 (#363)；文件后台归 #342 负责。

| 系列                       | 五份源码所显示的内容                                                                                                                                                                                                                                                                                                                                                        | 决策                                                                                                             |
| -------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| 状态栏                     | 计数器、保存状态、备注开关、视图按钮和缩放组共有五份手工构建的副本。                                                                                                                                                                                                                                                                                                        | 在 #390 中以 `pptx-ui-status-bar` **迁移**。                                                                     |
| 只读横幅                   | 全部五个绑定中的消息、仍然编辑、关闭和密码表单相同（`ReadOnlyBanner.tsx`、`ReadOnlyBanner.vue`、`readonly-banner.component.ts`、`ReadOnlyBanner.svelte`、`ui/read-only-banner.ts`）。偏差：React 和 Vanilla 没有 "Read-only recommended" 标题，Vanilla 使用关闭字形，角色在一个绑定中是 `alert`、在三个中是 `status`，并且只有 React 聚焦密码输入框。                       | **已迁移**：`pptx-ui-read-only-banner`。                                                                         |
| 粘贴选项工具栏             | 全部五个绑定中的四种格式条、锚定到画布的测量以及窗口级关闭都相同（`PasteOptionsToolbar.tsx/.vue/.svelte`、`paste-options-toolbar.component.ts`、`ui/paste-options-toolbar.ts`）。偏差：在条上按下可能在其点击生效之前就把它关闭，并且只有三个绑定会在选择格式后关闭它。                                                                                                     | **已迁移**：`pptx-ui-paste-options`。宿主仍然测量被粘贴的元素。                                                  |
| 兼容性提示                 | 全部五个绑定中的状态（提示、内边距）和回调（关闭一条、全部关闭）相同。偏差：Vue 和 Svelte 把堆叠限制为五条并带 "+N" 计数，其余三个显示全部，标记和图标也不同。                                                                                                                                                                                                              | **已迁移**：`pptx-ui-compat-toasts`，上限为五条，并通过 `compatToastStackStyle` 自行定位。                       |
| 对话框页脚                 | 每个对话框都在各自的模态外壳中手工构建页脚（React `useModalFocus`、Vue `ModalDialog.vue`、Angular `modal-dialog.component.ts`、Svelte `collab/components/ModalDialog.svelte`、Vanilla `parity-dialog-shell.ts`）。五个绑定中的按钮行形态相同。                                                                                                                              | 以 `pptx-ui-dialog-footer` **迁移**（不是外壳），被每个带按钮行的对话框采用 (#396，见下文逐对话框表格)。         |
| 移动端底部栏和顶部工具栏   | 相同的五个插槽和禁用规则（`buildBarActions`），以及相同的工具栏顺序（`MobileBottomBar.tsx/.vue`、`mobile-bottom-bar.component.ts`、`MobileActionSheets.svelte`、`ui/mobile-action-sheets.ts`；`MobileToolbar.tsx/.vue`、`mobile-toolbar.component.ts`、`MobileChrome.svelte`、`ui/mobile-toolbar.ts`）。偏差：标签键、评论徽标、打开面板的胶囊、仅编辑时的限制、AI 和共享。 | **已迁移**：`pptx-ui-mobile-bar` 和 `pptx-ui-mobile-toolbar`。                                                   |
| 演示工具栏和演讲者控制台条 | 二者都已经基于共享清单（`PRESENT_TOOLBAR_CONTROLS`、`PRESENTER_CONSOLE_CONTROLS`）渲染，id、顺序和标签键相同，但每个绑定都重新实现了标记、调色板、计时器和插槽状态规则。偏差：右键调色板仅在 React 和 Vue 中有，`aria-pressed` 仅在 Svelte 和 Vanilla 中有，“放大”的激活状态仅在 React、Vue 和 Angular 中有。                                                               | **已迁移**：`pptx-ui-present-toolbar` 和 `pptx-ui-presenter-console`。自动隐藏、批注和控制台快照补丁仍保持原生。 |
| 上下文菜单                 | 条目列表已经共享；渲染则在各绑定中为每个菜单手工构建（React `ContextMenu.tsx`、Vue `ContextMenu.vue`、Angular `editor-context-menu.component.ts`、Svelte `ElementContextMenu.svelte`、Vanilla `ui/element-context-menu.ts`，以及画布、缩略图和演示菜单）。没有任何绑定实现方向键、预输入查找或嵌套。                                                                        | 本次变更中未迁移（约 20 个适配器和 40 个测试文件）。具体范围见 #393。                                            |
| 标题栏和快速访问           | 插槽顺序相同，且 `pptx-ui-search` 已经共享。真实的偏差：Vanilla 按目录顺序渲染快速访问命令，而其他绑定写死保存、撤销和重做在前；工具提示、图标、保存限制和搜索限制各不相同；Svelte 从不传递 `disabledReason`。                                                                                                                                                              | 在偏差被决定之前保持原生；清单和计划见 #394。                                                                    |
| 备注工具栏和面板按钮       | 按钮集合相同，约定不同：Vue 使用另一个 i18n 命名空间，Svelte 在纯文本模式下隐藏格式并以不同方式排列缩进，链接 UI 在三个绑定中是弹出层、在两个中是 `window.prompt`，打印在 React 中是对话框、在其他地方是 iframe。                                                                                                                                                           | 在约定统一之前保持原生；计划见 #395。                                                                            |
| 幻灯片栏操作               | 功能不同：React 有“添加幻灯片”页脚，Vue 有复制/添加/删除行，Angular 有每个缩略图的移动按钮，Svelte 有节标题按钮；节菜单仅在 React 中是弹出层；浏览器菜单在 React 中有五项，在 Vue、Angular 和 Svelte 中有三项，在 Vanilla 中没有。幻灯片栏的其余部分是虚拟化、拖动重新排序和幻灯片内容渲染，而 #386 将这些排除在外。                                                        | 保持原生。需要先做出产品决策 (#397)。                                                                            |
| 检查器面板操作             | 约 170 个 React、144 个 Vue、124 个 Svelte 和 96 个 Vanilla 面板文件，各自通过绑定特有的编辑器和历史记录 API 提交，输入带防抖；叶子控件已经是共享的选择框和复选框。Svelte 和 Vanilla 还缺少其他绑定拥有的若干重置和清除操作（`image.resetCrop`、`media.resetTrim`、`chart.clearSeriesColor`、`slideBackground.clearBackground`）。                                          | 保持原生；一致性缺口在 #398 中跟踪。                                                                             |
| 文件后台                   | 归 #342 的原语工作负责。                                                                                                                                                                                                                                                                                                                                                    | 此处未触及。                                                                                                     |

### 新元素负责什么 {#what-the-new-elements-own}

八个新元素位于 `src/ui/src/pptx/web-components/`；它们的类型化 `state`、事件和行为记录在该目录的 README 中。它们遵循现有模式：开放的 shadow root、结构化的 `state` 属性、冒泡且 composed 的事件，以及对程序化更新不触发事件。适配器保留各绑定的公开 props、回调、id 和测试钩子（`pptx-readonly-*`、`pptx-compat-*`、`data-pptx-paste-options`、`data-pptx-present-toolbar`、`data-pptx-present-control`、`data-pptx-presenter-control`）。

行为变化，使五个绑定现在保持一致：

- 只读横幅：每个绑定都显示带锁和 `role="status"` 的 "Read-only recommended: message"，并在密码提示打开时聚焦密码输入框。
- 粘贴选项：在条内按下不再在点击生效之前将其关闭，且选择格式总是会将其关闭。在条内按 Escape 会关闭它。
- 兼容性提示：全部五个都把堆叠限制为五条并带 "+N" 计数，并始终提供 "Dismiss all"。
- 对话框页脚：取消和主要操作是同一个组件，已签名演示文稿的确认在所有地方都使用警告变体。`activateModalFocus` 现在会遍历开放的 shadow root，因此被困住的对话框仍可通过 Tab 到达其页脚。
- 移动端栏：一套标签键（Angular 过去对“插入”和“格式”使用另外两个键）、每个绑定中的评论徽标和打开面板的胶囊、每个绑定在查看模式下禁用仅编辑插槽，并且被点击的插槽保持其按下的颜色，因为悬停只对支持悬停的设备生效。
- 演示工具栏：每个切换项都暴露 `aria-pressed`，在每个绑定中右键“笔”或“荧光笔”都会打开其调色板，并且在粗指针下控件增大到 44px。宿主无法打开演示者视图时，省略演示者视图切换。
- 演讲者控制台：在每个绑定中，放大超过 100% 时“放大”显示为激活，并且一条规则（`presenterConsoleViewState`、`presenterConsoleAction`）取代了插槽状态逻辑的五份副本。保存这些副本的 Vue 和 Angular 辅助模块已被移除。

刻意不共享的部分：放映工具栏的自动隐藏包装器和批注模型、演示者快照补丁、每个对话框外壳和模态焦点陷阱、移动端面板，以及每一种文档变更。

### 验证 {#validation}

类型检查、`oxlint --deny-warnings` 和 `oxfmt` 在每个变更的文件上都通过。共享层和全部五个绑定的单元测试套件都通过（数量见 pull request）。适配器测试在每个绑定中覆盖每个已迁移系列的状态映射、意图路由、限制和回调替换；共享测试覆盖这些元素、它们的样式（44px 目标、强制颜色、令牌）、感知 shadow 的焦点陷阱以及插槽状态规则。

`e2e/pptx/chrome-controls-migration.spec.ts` 在全部五个绑定上运行，覆盖只读横幅（名称、密码焦点、错误和正确的密码、44px）、兼容性提示（关闭一条和全部、令牌、强制颜色、44px）、粘贴选项条（四种格式、选择、按键关闭）、对话框页脚（键盘取消和确定、令牌、强制颜色）、移动端栏（名称、顺序、44px、备注开关、令牌）和幻灯片放映工具栏（导航、工具按下状态、调色板选择、结束）。`present-chrome-parity`、`presenter-view-parity`、`paste-special`、`modify-password-*`、`compat-toast-clears-chrome`、`autosave-recovery-*`、`mobile-audit`、`mobile-notes`、`keyboard-shortcuts` 和其他放映规格都通过；两个探针（`present-chrome-parity`、`presenter-view-parity`）现在通过开放的 shadow root 读取控件。

状态栏和每个已迁移系列的前后对比截图位于 `docs/public/assets/ui-migration/non-ribbon-families/`（其 README 记录了演示文稿、视口和提交）。

## 每个对话框中的对话框页脚 (#396) {#dialog-footers-in-every-dialog-396}

`pptx-ui-dialog-footer` 现在在全部五个绑定中负责每个页脚为按钮行的对话框的操作行。外壳、遮罩、关闭和焦点陷阱仍归各绑定所有；回调、id 和测试钩子保持不变（诸如 `hf-apply-all` 的 `data-testid` 钩子通过 `actions[].testId` 传递）。

新增的元素变更，每一项都有单元测试：`align: 'start'` 把第三级操作固定在左边缘，`busy`（禁用、`aria-busy`、旋转指示器），`title`（工具提示）和 `danger` 变体。宿主为 `flex: 1 1 auto`，因此它会填满 flex 页脚，起始操作可以位于左侧。`modalFocusableElements` 和 `modalActiveElement` 为保留自有 Tab 循环的两个 Vanilla 对话框（SmartArt、幻灯片模板）暴露感知 shadow 的焦点遍历。

Enter 和 Escape 仍归各绑定的外壳所有：页脚只是按钮，因此在获得焦点的按钮上按 Enter 即激活它，而 Escape 属于外壳。没有任何对话框的页脚中带有绑定到 `<form>` 的表单提交或“不再显示”复选框，因此没有新增插槽。

图例：**M** 已迁移，**N** 保持原生，**-** 该绑定没有这个对话框。

| 对话框                                     | React | Vue | Angular | Svelte | Vanilla | 操作（顺序）                                    | 备注                                                                                                                     |
| ------------------------------------------ | ----- | --- | ------- | ------ | ------- | ----------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| 打印                                       | M     | M   | M       | M      | M       | 取消、打印（主要，打印图标）                    | Vue 和 React 把页数估算保留在页脚旁                                                                                      |
| 设置放映方式                               | M     | M   | M       | M      | M       | 取消、确定                                      |                                                                                                                          |
| 超链接                                     | M     | M   | M       | M      | M       | [删除链接，起始]、取消、应用或确定              | React 没有“删除链接”；Svelte 把主要按钮标为“确定”                                                                        |
| 文档属性                                   | M     | M   | M       | M      | M       | 取消、保存（未修改或只读时禁用）                | Vue 和 Angular 还有较小的“属性”对话框，以相同方式迁移                                                                    |
| 共享                                       | M     | M   | M       | M      | M       | 取消或关闭、开始或加入（实时时隐藏）            | 会话状态变化时重建操作列表                                                                                               |
| 广播                                       | M     | M   | M       | M      | M       | 关闭、开始（实时时隐藏）                        |                                                                                                                          |
| 公式编辑器                                 | M     | M   | M       | M      | M       | 取消、插入或更新（为空时禁用）                  | Vanilla 保留 Ctrl+Enter，它调用同一个提交                                                                                |
| OLE 编辑器                                 | M     | M   | M       | M      | M       | 替换文件（起始）、保存                          | 隐藏的文件输入仍位于该元素旁                                                                                             |
| 插入 SmartArt                              | M     | M   | M       | M      | M       | 取消、插入（未选择版式前禁用）                  | Vanilla 自己的 Tab 循环现在使用感知 shadow 的焦点遍历                                                                    |
| 幻灯片模板库                               | M     | M   | M       | M      | M       | 取消、插入（未选择图块前禁用）                  | Vanilla 的焦点变化相同                                                                                                   |
| 密码保护                                   | M     | M   | M       | M      | M       | [删除密码，危险，起始]、取消、设置或更新        |                                                                                                                          |
| 自定义放映                                 | -     | N   | -       | M      | M       | 取消、保存 (Svelte)；新建、取消、确定 (Vanilla) | React 在“设置放映方式”内编辑放映；Vue 和 Angular 使用带内联行命令和创建表单 (`type="submit"`) 的列表面板，因此没有页脚行 |
| 页眉和页脚                                 | M     | M   | M       | M      | M       | 全部应用、应用到当前（主要）                    | 保留 `hf-apply-all` 和 `hf-apply-current` 钩子                                                                           |
| 数字签名                                   | M     | -   | -       | M      | M       | 关闭或完成                                      | Vue 和 Angular 在面板而不是对话框中显示签名                                                                              |
| 字体嵌入                                   | M     | M   | M       | M      | M       | 完成（主要）                                    |                                                                                                                          |
| 加密文件提示                               | M     | -   | M       | -      | -       | 关闭                                            |                                                                                                                          |
| 日期和时间字段                             | M     | M   | M       | -      | -       | 取消、插入                                      | Svelte 和 Vanilla 不带选择器直接插入该字段                                                                               |
| 排练计时摘要                               | M     | M   | M       | M      | N       | 放弃、保存计时（主要）                          | Vanilla 在排练 HUD 栏中显示保存和放弃，那是工具栏而不是对话框                                                            |
| 备注打印预览                               | M     | -   | -       | -      | -       | 打印（主要，打印图标）                          |                                                                                                                          |
| 导出进度                                   | M     | M   | M       | M      | M       | 取消                                            |                                                                                                                          |
| 幻灯片大小缩放提示                         | N     | M   | -       | N      | -       | 最大化、确保适合（主要）                        | React 和 Svelte 把它渲染为检查器窗格中的紧凑内联卡片，而不是对话框                                                       |
| 选择性粘贴                                 | M     | M   | M       | M      | M       | 取消、确定                                      | 在 #386 中迁移                                                                                                           |
| 保留批注                                   | M     | M   | M       | M      | M       | 按提示而定                                      | 在 #386 中迁移                                                                                                           |
| 自动保存恢复                               | M     | M   | M       | M      | M       | 恢复、放弃                                      | 在 #386 中迁移                                                                                                           |
| 签名已移除                                 | M     | M   | M       | M      | M       | 取消、仍然编辑（警告）                          | 在 #386 中迁移                                                                                                           |
| 文件 > 选项（设置）                        | N     | N   | N       | N      | N       | 全部重置 (Vanilla)、取消、确定                  | 归 #342 负责；其页脚与整个对话框的实时重新标注绑定在一起                                                                 |
| 查找和替换                                 | N     | N   | N       | N      | N       | 查找、替换、全部替换、关闭                      | 由内联命令组成的停靠面板，其可访问名称与可见文本不同                                                                     |
| 设置背景格式                               | N     | N   | N       | N      | N       | 重置、关闭                                      | 带内联命令的停靠面板                                                                                                     |
| 比较、版本历史、辅助功能、快捷键、选择、AI | N     | N   | N       | N      | N       | 标题中的关闭                                    | 没有操作行的侧面板                                                                                                       |
| 移动端面板和菜单                           | N     | N   | N       | N      | N       | 无                                              | 底部面板，而不是对话框；没有页脚                                                                                         |

测试：共享层 (`chrome-controls`) 覆盖新的操作字段；每个绑定的适配器和对话框测试通过开放的 shadow root 查找页脚按钮（`dialog-footer.test-support`）；`e2e/pptx/dialog-footers-migration.spec.ts` 在每个绑定中打开九个对话框，检查页脚元素、主要操作位于最后、Escape（页脚保持完整）以及在获得焦点的“取消”上按 Enter。Escape 本身属于外壳，并不统一：React 的打印、超链接、设置放映方式和文档属性，Svelte 的同样几个对话框，以及 Vanilla 的文档属性和密码对话框都没有 Escape 处理程序（此前已存在，本次工作未改变，已作为后续跟踪）。前后对比截图位于 `docs/public/assets/ui-migration/dialog-footers/`。
