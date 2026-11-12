# Temas

La interfaz de Excel y los controles nativos de los frameworks comparten el sistema de diseño `--office-*` con Word,
PowerPoint y Visio. Instale el tema compartido una vez y use un ancestro común para aplicar el tema tanto
al lienzo como a los controles opcionales. Los controles nativos y los editores montados también lo instalan
de forma idempotente cuando se usan por separado.

```ts
import { installOfficeUiTheme } from 'ooxml-ui/theme';
installOfficeUiTheme();
```

```html
<section data-office-theme="dark" class="workbook-host">
	<!-- Render the native Toolbar, FormulaBar, SpreadsheetCanvas and SheetTabs here. -->
</section>
```

`data-office-theme` acepta `light`, `dark` y `auto`. Puede definirse en la raíz del documento o
en cualquier ancestro con ámbito. `auto` sigue el sistema operativo. Un lienzo con su `theme="auto"`
por defecto hereda la paleta de su ancestro. Un `theme="light"` o `theme="dark"` explícito
crea un ámbito local en ese editor, conservando la API de propiedades y atributos existente.
No cambia los controles hermanos del host ni una aplicación madre.

Los tokens canónicos se pueden sustituir en el mismo ámbito:

```css
.workbook-host {
	--office-accent: #1f9d63;
	--office-ring: #1f9d63;
	--office-background: #0f172a;
	--office-surface: #1e293b;
	--office-foreground: #f8fafc;
	--office-border: #64748b;
	--office-radius: 6px;
	--office-font: 'Segoe UI', sans-serif;
}
```

Los controles opcionales nativos se acogen a los estilos compartidos con los marcadores `data-office-block`. Sus
campos, botones, estados seleccionados, estados deshabilitados y anillos de foco siguen los mismos tokens.
La hoja de estilos no afecta a los controles personalizados del host que no tengan el marcador. Los cambios de tema actualizan el CSS
en su sitio, conservando el motor actual, la celda activa, el borrador, la selección y el campo con foco.
Los colores de las celdas del libro y el tema del documento no cambian; el papel de la hoja es blanco por defecto.
`--office-paper` y el antiguo `--xve-sheet-bg` son sustituciones de presentación.

## Sustituciones de tokens {#token-overrides}

`themeColors` sigue admitido para las sustituciones por editor. Cada clave se convierte en una propiedad personalizada `--xve-<kebab-key>` en el elemento host, con el mismo conjunto de tokens que usa el editor de Word (`--dve-*` allí), de modo que un único sistema de diseño puede controlar ambos:

| Clave                                  | Se usa para                                         |
| -------------------------------------- | --------------------------------------------------- |
| `background`, `foreground`             | Lienzo del editor y texto                           |
| `card`, `cardForeground`               | Cinta y paneles                                     |
| `popover`, `popoverForeground`         | Menús y desplegables                                |
| `primary`, `primaryForeground`         | Acento (barra de título, selección, pestaña activa) |
| `secondary`, `secondaryForeground`     | Botones secundarios                                 |
| `muted`, `mutedForeground`             | Encabezados, sugerencias, texto deshabilitado       |
| `accent`, `accentForeground`           | Estados al pasar el cursor y al pulsar              |
| `destructive`, `destructiveForeground` | Errores                                             |
| `border`, `input`, `ring`              | Líneas, campos y anillos de foco                    |
| `radius`                               | Radio de las esquinas                               |

```ts
editor.themeColors = { primary: '#1f9d63', ring: '#1f9d63', radius: '6px' };
```

También puede definir las propiedades personalizadas en CSS:

```css
xlsx-editor {
	--xve-primary: #1f9d63;
}
```

## Seguir el tema de un sitio {#following-a-site-theme}

La demo sigue la clave `vitepress-theme-appearance` de VitePress en `localStorage` y define
`data-office-theme` en la raíz del documento. Enlace ese mismo atributo a la configuración de su aplicación para que los controles compuestos y los editores cambien juntos. El editor nunca lee
`localStorage` por sí mismo.

Las sustituciones CSS heredadas `--xve-*` siguen admitidas en el editor o en un ancestro. Tienen prioridad sobre el token canónico correspondiente cuando se proporcionan. Los primitivos compartidos internos
reciben esos valores ya resueltos sin volver a declarar la paleta canónica en el host del editor. Los objetos de preajuste claro/oscuro existentes y las claves de `themeColors` conservan su forma pública.
