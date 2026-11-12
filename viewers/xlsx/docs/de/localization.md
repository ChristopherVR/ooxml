# Lokalisierung

Die Oberfläche des Editors ist für fünf Sprachen gebaut und wird über das Attribut `locale`, die gleichnamige Eigenschaft oder die Prop des Bindings gewählt:

| Wert    | Sprache                  |
| ------- | ------------------------ |
| `en`    | Englisch (Standard)      |
| `fr`    | Französisch              |
| `de`    | Deutsch                  |
| `es`    | Spanisch                 |
| `zh-CN` | Vereinfachtes Chinesisch |

Jedes BCP-47-Tag wird auf eine dieser Sprachen abgebildet (`normalizeEditorLocale`): `de-DE` wird zu `de`, `es-MX` wird zu `es`, `zh-Hans` wird zu `zh-CN`. Traditionelles Chinesisch (`zh-TW`, `zh-Hant`) und nicht unterstützte Sprachen fallen auf Englisch zurück.

```ts
editor.locale = navigator.language;
```

Die Lokalisierung betrifft nur den Oberflächentext: Menübandbeschriftungen, Dialoge, Menüs, die Statusleiste und Meldungen. Sie übersetzt nie den Inhalt einer Arbeitsmappe. Funktionsnamen in Formeln bleiben auf Englisch (so, wie sie in der Datei gespeichert sind), und Zahlen- und Datumsformate werden so dargestellt, wie die Arbeitsmappe sie festlegt, nicht nach der Oberflächensprache.

Die Demo nimmt `?locale=` in ihrer URL entgegen und bietet eine Sprachauswahl an.
