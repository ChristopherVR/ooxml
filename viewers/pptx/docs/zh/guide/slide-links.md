# 幻灯片链接 {#slide-links}

指向幻灯片的文本片段和形状超链接，在只读模式、预览模式以及幻灯片放映期间，都会在查看器内部导航。内部 `ppaction://` 链接绝不会打开浏览器标签页。编辑时，请使用 Ctrl+单击（macOS 上为 Cmd+单击）来跟随内部幻灯片链接；普通单击仍用于选择或编辑元素。

只读和预览查看器会隐藏琥珀色的操作指示器及其编辑提示。编辑查看器则继续显示它们。

## 宿主回调 {#host-callback}

每个绑定都提供一个回调，在跟随文本片段链接或形状链接之前调用。返回 `false` 可取消默认导航，包括外部导航。

| 绑定         | 回调                                           |
| ------------ | ---------------------------------------------- |
| Vanilla 选项 | `onHyperlinkClick`                             |
| React 属性   | `onHyperlinkClick`                             |
| Vue 回调属性 | `onHyperlinkClick`（或 `:on-hyperlink-click`） |
| Angular 输入 | `[onHyperlinkClick]`                           |
| Svelte 属性  | `onhyperlinkclick`                             |

回调接收来自 `ooxml-ui/pptx` 的 `HyperlinkClickEvent`：外部 `url`、内部 `action`、解析成功时从零开始的 `targetSlideIndex`，以及可用时形状的 `elementId`。没有值的字段为 undefined。

```ts
const viewer = createPptxViewer(container, {
	onHyperlinkClick: (link) => {
		if (link.targetSlideIndex !== undefined) {
			selectSlideInHost(link.targetSlideIndex);
			return false;
		}
	},
});
```

没有回调时，查看器会正常跟随链接。外部链接保留演示文稿的目标框架、URL 安全检查以及信任中心确认设置。
