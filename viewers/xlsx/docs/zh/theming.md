# 主题

Excel 界面框架与原生框架控件共享 `--office-*` 设计系统，该系统也被 Word、PowerPoint 和 Visio 使用。只需安装一次共享主题，然后使用一个共同的祖先元素，即可同时为画布和可选控件设置主题。原生控件和已挂载的编辑器在单独使用时，也会以幂等方式安装该主题。

```ts
import { installOfficeUiTheme } from 'ooxml-ui/theme';
installOfficeUiTheme();
```

```html
<section data-office-theme="dark" class="workbook-host">
	<!-- Render the native Toolbar, FormulaBar, SpreadsheetCanvas and SheetTabs here. -->
</section>
```

`data-office-theme` 接受 `light`、`dark` 和 `auto`。它可以设置在文档根元素或任何作用域祖先元素上。`auto` 跟随操作系统。使用默认 `theme="auto"` 的画布会继承其祖先的配色。显式设置 `theme="light"` 或 `theme="dark"` 会在该编辑器上创建本地作用域，同时保留现有的属性和特性 API。它不会改变同级的宿主控件或父应用程序。

可以在同一作用域上覆盖标准令牌：

```css
.workbook-host {
	--office-accent: #1f9d63;
	--office-ring: #1f9d63;
	--office-background: #0f172a;
	--office-surface: #1e293b;
	--office-foreground: #f8fafc;
	--office-border: #64748b;
	--office-radius: 6px;
	--office-font: 'Segoe UI', sans-serif;
}
```

原生可选控件通过 `data-office-block` 标记选择加入共享样式。它们的输入框、按钮、选中状态、禁用状态和焦点环都遵循相同的令牌。样式表不会作用于未标记的自定义宿主控件。主题更改会就地更新 CSS，并保留当前引擎、活动单元格、草稿、选区和已获得焦点的输入框。工作簿单元格颜色和文档主题保持不变；工作表背景默认为白色。`--office-paper` 和旧版 `--xve-sheet-bg` 是呈现层的覆盖项。

## 令牌覆盖 {#token-overrides}

`themeColors` 仍然支持针对每个编辑器进行覆盖。每个键都会在宿主元素上成为一个 `--xve-<kebab-key>` 自定义属性，其令牌集与 Word 编辑器所用的相同（在 Word 中为 `--dve-*`），因此同一套设计系统可以同时驱动两者：

| 键                                     | 用途                               |
| -------------------------------------- | ---------------------------------- |
| `background`、`foreground`             | 编辑器画布和文本                   |
| `card`、`cardForeground`               | 功能区和面板                       |
| `popover`、`popoverForeground`         | 菜单和下拉列表                     |
| `primary`、`primaryForeground`         | 强调色（标题栏、选区、活动选项卡） |
| `secondary`、`secondaryForeground`     | 次要按钮                           |
| `muted`、`mutedForeground`             | 标题、提示和禁用文本               |
| `accent`、`accentForeground`           | 悬停和按下状态                     |
| `destructive`、`destructiveForeground` | 错误                               |
| `border`、`input`、`ring`              | 线条、输入框和焦点环               |
| `radius`                               | 圆角半径                           |

```ts
editor.themeColors = { primary: '#1f9d63', ring: '#1f9d63', radius: '6px' };
```

也可以在 CSS 中设置自定义属性：

```css
xlsx-editor {
	--xve-primary: #1f9d63;
}
```

## 跟随站点主题 {#following-a-site-theme}

演示会跟随 `localStorage` 中的 VitePress `vitepress-theme-appearance` 键，并在文档根元素上设置 `data-office-theme`。请将同一属性绑定到你的应用程序设置，以便组合的控件和编辑器同步变化。编辑器本身从不读取 `localStorage`。

旧版 `--xve-*` CSS 覆盖仍可在编辑器或其祖先元素上使用。若提供了对应的值，它们的优先级高于相应的标准令牌。内部共享的基础组件会接收这些已解析的值，而无需在编辑器宿主上重新声明标准配色。现有的浅色/深色预设对象和 `themeColors` 键保持其公开结构不变。
