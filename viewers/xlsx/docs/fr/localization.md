# Localisation

L’interface de l’éditeur est conçue pour cinq langues, choisies avec l’attribut `locale`, la propriété ou la prop de liaison :

| Valeur  | Langue               |
| ------- | -------------------- |
| `en`    | Anglais (par défaut) |
| `fr`    | Français             |
| `de`    | Allemand             |
| `es`    | Espagnol             |
| `zh-CN` | Chinois simplifié    |

Toute balise BCP 47 correspond à l’une d’elles (`normalizeEditorLocale`) : `de-DE` devient `de`, `es-MX` devient `es`, `zh-Hans` devient `zh-CN`. Le chinois traditionnel (`zh-TW`, `zh-Hant`) et les langues non prises en charge reviennent à l’anglais.

```ts
editor.locale = navigator.language;
```

La localisation ne couvre que le texte de l’interface : libellés du ruban, boîtes de dialogue, menus, barre d’état et messages. Elle ne traduit jamais le contenu du classeur. Les noms de fonctions dans les formules restent en anglais (tels qu’ils sont stockés dans le fichier), et les formats de nombres et de dates s’affichent selon la définition du classeur, et non selon la langue de l’interface.

La démo accepte `?locale=` dans son URL et propose un sélecteur de langue.
