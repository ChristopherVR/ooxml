# Localización

La interfaz del editor está preparada para cinco idiomas, que se seleccionan con el atributo `locale`, la propiedad o la prop del binding:

| Valor   | Idioma               |
| ------- | -------------------- |
| `en`    | Inglés (por defecto) |
| `fr`    | Francés              |
| `de`    | Alemán               |
| `es`    | Español              |
| `zh-CN` | Chino simplificado   |

Cualquier etiqueta BCP 47 se asigna a uno de ellos (`normalizeEditorLocale`): `de-DE` pasa a `de`, `es-MX` pasa a `es` y `zh-Hans` pasa a `zh-CN`. El chino tradicional (`zh-TW`, `zh-Hant`) y los idiomas no admitidos vuelven al inglés.

```ts
editor.locale = navigator.language;
```

La localización abarca solo el texto de la interfaz: etiquetas de la cinta, diálogos, menús, la barra de estado y los mensajes. Nunca traduce el contenido del libro. Los nombres de las funciones en las fórmulas se quedan en inglés (tal como están guardados en el archivo), y los formatos de número y de fecha se muestran según los define el libro, no según el idioma de la interfaz.

La demo acepta `?locale=` en su URL y ofrece un selector de idioma.
