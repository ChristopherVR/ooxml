# 主题

查看器自行绘制其界面（功能区、后台视图、状态栏、形状窗口），这些界面来自共享的 `ooxml-ui` 控件。元素从 CSS 自定义属性中读取颜色，因此宿主可以在自己的样式表中重新设置其外观。

## 共享的 `--office-*` 令牌 {#shared-office-tokens}

元素遵循共享的 Office 主题。连接时，它会安装 `ooxml-ui` 主题（`installOfficeUiTheme`，可重复调用而不产生副作用），其绘制的每种颜色都取自 `--office-*` 自定义属性，因此 `--office-*` 令牌是主要的主题定制方式。可以在 `:root`、容器或元素本身上设置它们：

| 令牌                         | 作用                     |
| ---------------------------- | ------------------------ |
| `--office-background`        | 功能区、面板和对话框表面 |
| `--office-surface`           | 窗口背景、次要表面       |
| `--office-foreground`        | 主要文本                 |
| `--office-muted-foreground`  | 次要文本                 |
| `--office-border`            | 边框                     |
| `--office-accent`            | 强调色                   |
| `--office-accent-foreground` | 强调色上的文本           |
| `--office-ring`              | 焦点环                   |
| `--office-danger`            | 破坏性操作和错误状态     |
| `--office-shadow`            | 页面阴影                 |

```css
visio-viewer {
	--office-accent: #c2431f;
	--office-accent-foreground: #ffffff;
}
```

浅色和深色外观跟随套件主题：查看器默认为浅色，系统偏好深色时则为深色；文档根元素上的 `data-office-theme="light" | "dark"` 属性可强制使用某一种。元素上没有 `theme` 属性；页面上用于切换其他 Office 元素主题的同一开关，也会切换查看器的主题。

默认强调色也因此发生了变化。此前查看器硬编码为 `#3955a3`；现在它读取 `--office-accent`（浅色为 `#2563eb`，深色为 `#6366f1`）。若宿主希望保留原来的颜色，请设置 `--office-accent: #3955a3` 或下文的 `--vv-accent` 覆盖项。

## 可选的 `--vv-*` 覆盖项 {#optional-vv-overrides}

早期的 `--vv-*` 令牌仍然有效。每个令牌一旦设置，就会优先于它所别名的 `--office-*` 令牌；未设置时，查看器则回退到共享令牌。若只想重新设置本查看器的外观而不影响页面的其他部分，可以使用它们。

| 令牌               | 回退到                       | 作用                         |
| ------------------ | ---------------------------- | ---------------------------- |
| `--vv-font`        | `--office-font`              | 字体系列                     |
| `--vv-background`  | `--office-surface`           | 窗口背景                     |
| `--vv-surface`     | `--office-background`        | 功能区和面板表面             |
| `--vv-secondary`   | `--office-surface`           | 次要表面                     |
| `--vv-shadow`      | `--office-shadow`            | 页面阴影                     |
| `--vv-ink`         | `--office-foreground`        | 主要文本                     |
| `--vv-muted`       | `--office-muted-foreground`  | 次要文本                     |
| `--vv-border`      | `--office-border`            | 边框                         |
| `--vv-accent`      | `--office-accent`            | 强调色                       |
| `--vv-accent-soft` | 强调色的 10%                 | 柔和的强调表面（悬停、选中） |
| `--vv-accent-ink`  | `--office-accent-foreground` | 强调色上的文本               |
| `--vv-focus`       | `--office-ring`              | 焦点环                       |
| `--vv-danger`      | `--office-danger`            | 破坏性操作和错误状态         |

演示环境会在文档根元素上设置 `data-theme="light" | "dark"`，并在 `demos/visio/demo/workspace.css` 中为其自身的页面外壳（起始界面、工作区外壳）提供相应的 `--vv-*` 值；该文件仍然有效，因为 `--vv-*` 覆盖项依然生效。

## 跟随文档站点主题 {#following-the-docs-site-theme}

本站将访问者的外观选择存储在 `localStorage` 中，键名为 VitePress 的 `vitepress-theme-appearance`（值为 `light`、`dark`，或无值，表示跟随系统）。所有查看器站点共享 `christophervr.github.io` 源，演示环境（`demos/visio/demo/suite-theme.js`）会在启动时读取同一个键，并通过 `storage` 事件实时更新。因此，在本站切换主题会重新设置嵌入查看器的样式，无需重新加载，也不会替换其图纸。

```ts
const key = 'vitepress-theme-appearance';
const resolve = (value: string | null) =>
	value === 'dark' || value === 'light'
		? value
		: matchMedia('(prefers-color-scheme: dark)').matches
			? 'dark'
			: 'light';

document.documentElement.dataset.theme = resolve(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null)
		document.documentElement.dataset.theme = resolve(event.newValue);
});
```

## 状态栏主题 {#status-bar-theme}

与套件中的其他编辑器一样，状态栏默认为中性色。桌面版 Visio 的“彩色”Office 主题则会用产品颜色填充状态栏。有三种方式可以选择：

- **用户**：文件 > 选项 > 常规 > Office 主题（彩色、白色或黑色，其中黑色为深色主题）。该选择保存在设备上。
- **宿主**：在 `<visio-viewer>` 上设置 `status-bar` 属性（`colorful` 或 `neutral`），或设置 `statusBar` 属性。宿主的属性优先于已保存的选择。
- **任意颜色**：`--vv-status-background` 和 `--vv-status-ink` 可在任一主题下重新设置状态栏的颜色。

```html
<visio-viewer status-bar="colorful"></visio-viewer>
<visio-viewer style="--vv-status-background: #1b5e20; --vv-status-ink: #fff"></visio-viewer>
```
