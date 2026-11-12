# Folienlinks

Hyperlinks in Text-Runs und Formen, die auf eine Folie verweisen, navigieren im Viewer innerhalb des Dokuments, und zwar im Lese- und Vorschaumodus sowie während einer Bildschirmpräsentation. Interne `ppaction://`-Links öffnen nie einen Browser-Tab. Beim Bearbeiten folgen Sie einem internen Folienlink mit Strg+Klick (Cmd+Klick unter macOS); ein normaler Klick wählt das Element weiterhin aus oder bearbeitet es.

Schreibgeschützte Viewer und Vorschau-Viewer blenden den Aktionsindikator in Bernsteinfarbe und seinen Hinweis zur Bearbeitung aus. Bearbeitende Viewer zeigen beides weiterhin an.

## Host-Callback {#host-callback}

Jede Bindung bietet einen Callback an, der vor dem Folgen eines Text-Run- oder Formlinks aufgerufen wird. Geben Sie `false` zurück, um die Standardnavigation abzubrechen, einschließlich der Navigation zu externen Zielen.

| Bindung           | Callback                                        |
| ----------------- | ----------------------------------------------- |
| Vanilla-Optionen  | `onHyperlinkClick`                              |
| React-Prop        | `onHyperlinkClick`                              |
| Vue-Callback-Prop | `onHyperlinkClick` (oder `:on-hyperlink-click`) |
| Angular-Eingabe   | `[onHyperlinkClick]`                            |
| Svelte-Prop       | `onhyperlinkclick`                              |

Der Callback erhält das `HyperlinkClickEvent` aus `ooxml-ui/pptx`: die externe `url`, die interne `action`, einen nullbasierten `targetSlideIndex`, sofern er aufgelöst werden konnte, und die `elementId` der Form, sofern verfügbar. Felder ohne Wert sind undefined.

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

Ohne Callback folgt der Viewer Links auf die übliche Weise. Externe Links behalten das Zielfenster der Präsentation sowie die URL-Sicherheitsprüfungen und die Einstellung der Bestätigung im Trust Center bei.
