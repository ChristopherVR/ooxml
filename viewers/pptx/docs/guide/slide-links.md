# Slide links

Text-run and shape hyperlinks to a slide navigate within the viewer in read-only
and preview mode, as well as during a slide show. Internal `ppaction://` links
never open a browser tab. While editing, use Ctrl+Click (Cmd+Click on macOS) to
follow an internal slide link; ordinary clicks keep selecting or editing the element.

Read-only and preview viewers hide the amber action indicator and its authoring
tooltip. Editing viewers continue to show them.

## Host callback

Each binding offers a callback before following either a text-run or shape link.
Return `false` to cancel the default navigation, including external navigation.

| Binding           | Callback                                      |
| ----------------- | --------------------------------------------- |
| Vanilla options   | `onHyperlinkClick`                            |
| React prop        | `onHyperlinkClick`                            |
| Vue callback prop | `onHyperlinkClick` (or `:on-hyperlink-click`) |
| Angular input     | `[onHyperlinkClick]`                          |
| Svelte prop       | `onhyperlinkclick`                            |

The callback receives `HyperlinkClickEvent` from `ooxml-ui/pptx`: the external
`url`, the internal `action`, a zero-based `targetSlideIndex` when resolved, and
the shape's `elementId` when available. Fields without a value are undefined.

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

Without a callback, the viewer follows links normally. External links retain the
deck's target frame, URL safety checks and Trust Center confirmation setting.
