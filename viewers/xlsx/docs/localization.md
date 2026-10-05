# Localization

The editor interface is built for five languages, selected with the `locale` attribute, property or binding prop:

| Value   | Language           |
| ------- | ------------------ |
| `en`    | English (default)  |
| `fr`    | French             |
| `de`    | German             |
| `es`    | Spanish            |
| `zh-CN` | Simplified Chinese |

Any BCP 47 tag maps onto one of them (`normalizeEditorLocale`): `de-DE` becomes `de`, `es-MX` becomes `es`, `zh-Hans` becomes `zh-CN`. Traditional Chinese (`zh-TW`, `zh-Hant`) and unsupported languages fall back to English.

```ts
editor.locale = navigator.language;
```

Localization covers interface text only: ribbon labels, dialogs, menus, the status bar and messages. It never translates workbook content. Function names in formulas stay in English (as stored in the file), and number and date formats render as the workbook defines them, not by the interface locale.

The demo takes `?locale=` in its URL and offers a language picker.
