# 主题

`<teams-app>` 的所有颜色、间距值和圆角都来自共享的 `ooxml-ui` 令牌（`var(--office-...)`）。自定义属性可以穿透 Shadow DOM 边界，因此在 `:root`（或某个容器）上设置一条规则，即可为整个应用设置主题。

## 浅色与深色 {#light-and-dark}

令牌默认跟随操作系统的配色方案。若要强制使用某种配色方案，请在根元素上设置 `data-office-theme`：

```ts
document.documentElement.setAttribute('data-office-theme', 'dark'); // or 'light'
document.documentElement.removeAttribute('data-office-theme'); // follow the OS again
```

## 覆盖令牌 {#token-overrides}

在希望生效的位置覆盖任意令牌：

```css
:root {
	--office-accent: #0e8f8f;
	--office-teams-brand: #0e8f8f;
	--office-teams-rail: #14171b;
}
```

Teams 专用令牌以 `--office-teams-` 为前缀（品牌色、侧栏、面板、气泡、分隔线和次要文字）。该应用没有自己的 `theme` 属性：主题通过这些令牌来设置。本站文档的强调色是有意选用的中性青色；OpenTeams 不借用 Microsoft 的品牌色。

## 跟随文档站点主题 {#following-the-docs-site-theme}

本站将访问者的外观选择保存在 `localStorage` 中，键名为 VitePress 的 `vitepress-theme-appearance`（取值 `light`、`dark` 或 `auto`）。所有查看器站点共享 `christophervr.github.io` 源，而本仓库中的演示应用（仓库根目录的 `demos/teams/theme-sync.ts`）会在启动时以及 `storage` 事件发生时应用该设置，因此在本站切换主题会即时更新嵌入的演示，无需重新加载：

```ts
const key = 'vitepress-theme-appearance';
const apply = (value: string | null) => {
	const root = document.documentElement;
	if (value === 'light' || value === 'dark') root.setAttribute('data-office-theme', value);
	else root.removeAttribute('data-office-theme');
};
apply(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null) apply(event.newValue);
});
```
