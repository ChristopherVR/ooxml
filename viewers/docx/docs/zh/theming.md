# 主题

编辑器为自身的界面（标题栏、功能区、面板、状态栏、对话框）设置主题。文档的页面始终保持白色，与 Word 的页面一致。

## 浅色、深色与自动 {#light-dark-and-auto}

将 `theme` 属性（特性 `theme`）设置为 `light`、`dark` 或 `auto`。无法识别的值会回退为 `auto`：它会继承最近的共享 Office 主题作用域；如果没有设置作用域，则跟随操作系统的配色方案。

```ts
editor.theme = 'dark';
```

## 共享的 Office 令牌 {#shared-office-tokens}

Word 与 Excel、Visio、PowerPoint 以及共享控件使用相同的 `--office-*` 令牌。在公共祖先元素上设置 `data-office-theme="light"`、`"dark"` 或 `"auto"`，即可为整个组合设置主题。其模式和自定义令牌会实时更新，包括跨越编辑器的 Shadow Root。原生的 React、Vue、Angular、Svelte 和 Solid 工具栏组件同样采用这些共享样式。

```html
<section data-office-theme="dark" class="office-host">
	<!-- native canvas and toolbar components -->
</section>
<style>
	.office-host {
		--office-accent: #9b59b6;
		--office-ring: #b37cd0;
		--office-radius: 6px;
		--office-font: system-ui, sans-serif;
	}
</style>
```

窗口外观使用 `--office-background`，面板和工作区使用 `--office-surface`，界面文字使用 `--office-foreground`，轮廓线使用 `--office-border`。内置控件还共享 Office 的字体排版、间距、焦点和禁用状态令牌。宿主可以用自己的样式覆盖这些优先级较低的默认值。文档的文本格式和序列化的 Office 主题保持不变。在深色模式下，页面默认仍为白色；`--office-paper` 以及现有的 `--dve-page-color` 提供显式的呈现覆盖。

显式设置编辑器的 `theme="light"` 或 `theme="dark"` 会为其画布创建一个局部作用域。若要同时为同级的原生控件设置主题，请使用公共祖先；编辑器不会改变其宿主应用的主题。将编辑器恢复为 `auto` 后，会恢复继承的作用域行为。

## 旧版令牌覆盖 {#legacy-token-overrides}

`--dve-*` 覆盖仍然受支持，并优先于其规范回退值。`themeColors` 接受部分 `EditorTheme`，并将每个键作为 `--dve-*` 自定义属性应用到宿主上，在浅色和深色模式下均有效。值可以是任何 CSS 颜色。

```ts
editor.themeColors = {
	primary: '#c2431f',
	primaryForeground: '#ffffff',
	radius: '6px',
};
```

| 令牌                                   | 作用                           |
| -------------------------------------- | ------------------------------ |
| `background`、`foreground`             | 应用外壳（标题栏、窗口）       |
| `card`、`cardForeground`               | 功能区、面板和状态栏的表面     |
| `popover`、`popoverForeground`         | 菜单、对话框和其他浮动表面     |
| `primary`、`primaryForeground`         | 品牌色和主要操作颜色           |
| `secondary`、`secondaryForeground`     | 次要表面                       |
| `muted`、`mutedForeground`             | 弱化的表面，也是页面背后的区域 |
| `accent`、`accentForeground`           | 悬停和选中的高亮               |
| `destructive`、`destructiveForeground` | 破坏性操作                     |
| `border`、`input`、`ring`              | 边框、输入框轮廓、焦点环       |
| `radius`                               | 基础圆角半径，例如 `4px`       |

传入 `undefined` 可清除这些覆盖。

## 跟随文档网站主题 {#following-the-docs-site-theme}

本站将访问者的外观选择存储在 `localStorage` 中，使用 VitePress 的键 `vitepress-theme-appearance`（取值为 `light`、`dark` 或 `auto`）。所有查看器站点共享 `christophervr.github.io` 源，演示应用在启动时读取同一个键，并通过 `storage` 事件实时更新。因此在本站切换主题，会重新设置嵌入编辑器的样式，而无需重新加载。简而言之：

```ts
const key = 'vitepress-theme-appearance';
const resolve = (value: string | null) =>
	value === 'dark' || value === 'light'
		? value
		: matchMedia('(prefers-color-scheme: dark)').matches
			? 'dark'
			: 'light';

editor.theme = resolve(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null) editor.theme = resolve(event.newValue);
});
```

宿主应用可以自由使用自己的键：该元素只需要 `theme` 属性即可。
