---
title: Temas
description: 'El sistema ViewerTheme común a todos los enlaces: la forma del objeto de tema, cómo themeToCssVars genera las propiedades personalizadas --pptx-*, los temas predefinidos, la aplicación por enlace y el selector Archivo > Opciones > Apariencia.'
---

# Temas

La interfaz del visualizador (barra de herramientas, cinta, cuadros de diálogo, backstage) se tema mediante un objeto `ViewerTheme`: un conjunto de propiedades personalizadas CSS (`--pptx-*`) aplicadas a la raíz del visualizador. Esto es totalmente independiente del esquema de colores y de las fuentes propios de la presentación de OOXML (la galería "Temas" de la pestaña Diseño edita el propio documento `.pptx`); `ViewerTheme` solo afecta a la interfaz de la aplicación.

El sistema de temas no depende de ningún framework. Los tipos, valores por defecto, predefinidos y funciones auxiliares que se describen a continuación se implementan una sola vez (en la API pública de `ooxml-ui/pptx`) y los reexportan de forma idéntica todos los enlaces publicados: `pptx-react-viewer`, `pptx-vue-viewer`, `pptx-angular-viewer`, `pptx-svelte-viewer` y `pptx-vanilla-viewer`.

## Apariencia compartida de Office {#shared-office-appearance}

Los cinco enlaces heredan los mismos tokens `--office-*` que Word, Excel y Visio. Ponga `data-office-theme="light"`, `"dark"` o `"auto"` en la raíz del documento o en un elemento contenedor. Los modos aplicados a un ámbito anulan la paleta heredada, y `auto` sigue al sistema operativo cuando cambia. Los tokens personalizados en el mismo ámbito anulan la paleta seleccionada:

```html
<section data-office-theme="dark" style="--office-accent: #a78bfa; --office-font: Georgia, serif">
	<!-- Mount a viewer here, without an explicit ViewerTheme palette. -->
</section>
```

La barra de herramientas, los controles nativos y los controles de shadow DOM heredan el ámbito más cercano. Entre los tokens útiles están `--office-background`, `--office-surface`, `--office-foreground`, `--office-border`, `--office-accent`, `--office-ring` y `--office-font`. Las salidas existentes de `ViewerTheme`, `themeToCssVars()` y `defaultCssVars()` siguen siendo compatibles. Las entradas heredadas explícitas siguen prevaleciendo sobre los valores de reserva de los tokens canónicos, incluidas las paletas parciales y las dimensiones de los controles, como `--pptx-switch-width`.

La apariencia de la interfaz no modifica los colores, las fuentes ni el tema del documento de una presentación.

## La forma de `ViewerTheme` {#the-viewertheme-shape}

```ts
interface ViewerTheme {
	/** Semantic UI colors. Each key maps to a `--pptx-<key>` custom property. */
	colors?: Partial<ViewerThemeColors>;
	/** Base border-radius value (e.g. "0.5rem", "8px"). */
	radius?: string;
	/** Escape hatch: arbitrary CSS custom properties set on the viewer root. Keys include the `--` prefix. */
	cssVars?: Record<string, string>;
}
```

Todos los campos son opcionales; los valores no definidos heredan los tokens compartidos de la interfaz de Office en el ámbito del visualizador. Los colores explícitos de `ViewerTheme` y las variables `--pptx-*` heredadas tienen prioridad. `ViewerThemeColors` tiene 19 tokens semánticos, con nombres según la convención de shadcn/ui. Se acepta cualquier cadena de color CSS válida (hex, `rgb()`, `hsl()`, `oklch()`, colores con nombre):

| Tokens                                 | Función                                         |
| -------------------------------------- | ----------------------------------------------- |
| `background`, `foreground`             | Fondo de la raíz y texto predeterminado         |
| `card`, `cardForeground`               | Superficies de tarjetas y paneles               |
| `popover`, `popoverForeground`         | Ventanas emergentes y desplegables              |
| `primary`, `primaryForeground`         | Acciones principales (botones, estados activos) |
| `secondary`, `secondaryForeground`     | Acciones secundarias                            |
| `muted`, `mutedForeground`             | Superficies discretas y texto secundario        |
| `accent`, `accentForeground`           | Superficies de resaltado al pasar el cursor     |
| `destructive`, `destructiveForeground` | Acciones de peligro o eliminación               |
| `border`, `input`, `ring`              | Bordes, bordes de campos y anillo de foco       |

Consulte la [página de temas de React](/react/theming#viewertheme-and-viewerthemecolors) para ver la tabla completa, token por token, con el nombre exacto de la variable CSS de cada uno.

## Cómo `themeToCssVars` genera las variables `--pptx-*` {#how-themetocssvars-produces-pptx-variables}

`themeToCssVars(theme, omitDefaults = false)` convierte un `ViewerTheme` en un `Record<string, string>` plano de propiedades personalizadas CSS, listo para aplicarse a la raíz del visualizador como estilo en línea. Cada enlace la llama internamente cuando usted pasa un tema; también se exporta para crear sus propias herramientas.

- Cada clave de `colors` se convierte en `--pptx-<clave-en-kebab-case>`: `primaryForeground: '#fff'` genera `--pptx-primary-foreground: #fff`.
- Cada color se refleja también en el token semántico de Tailwind correspondiente (`--color-primary-foreground`), para compatibilidad con los hosts de Tailwind CSS v4. El visualizador usa `@theme inline`, de modo que los tokens de Office heredados se resuelven en el componente que los consume.
- `radius` se convierte en `--pptx-radius`, más los valores derivados `--radius-sm` / `--radius-md` / `--radius-lg` / `--radius-xl` (de `calc(r - 4px)` a `calc(r + 4px)`).
- Las entradas de `cssVars` pasan tal cual.
- Con `omitDefaults: true`, se omiten los valores iguales a los valores predeterminados integrados.

```ts
import { themeToCssVars, defaultCssVars } from 'pptx-react-viewer';

themeToCssVars({ colors: { primary: '#6366f1' }, radius: '0.75rem' });
// {
//   '--pptx-primary': '#6366f1', '--color-primary': '#6366f1',
//   '--pptx-radius': '0.75rem',
//   '--radius-sm': 'calc(0.75rem - 4px)', ... '--radius-xl': 'calc(0.75rem + 4px)',
// }

defaultCssVars();
// The complete set of --pptx-* properties with the built-in dark defaults,
// for generating a full fallback stylesheet.
```

## Temas y predefinidos integrados {#built-in-themes-and-presets}

| Exportación                                    | Paleta                                                                                                                                                  |
| ---------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `defaultThemeColors` + `defaultRadius`         | La interfaz oscura integrada (escala de grises de Tailwind, primario índigo, radio `0.5rem`). Disponible explícitamente a través de `defaultCssVars()`. |
| `vermilionLightTheme` / `vermilionLightColors` | Paleta clara cálida tipo "papel" con el acento bermellón que usa esta documentación.                                                                    |
| `vermilionDarkTheme` / `vermilionDarkColors`   | Paleta oscura atenuada tipo "sala del presentador" con el mismo acento.                                                                                 |
| `vermilionRadius`                              | `'0.375rem'`, el radio que usan ambos predefinidos vermilion.                                                                                           |

Los dos predefinidos vermilion son objetos `ViewerTheme` completos (los 19 tokens más el radio), así que sustituyen por completo a los valores oscuros por defecto. Las paletas `*Colors` sin procesar se exportan junto a ellos para derivar variantes:

```ts
import { vermilionDarkColors, vermilionRadius } from 'pptx-react-viewer';
import type { ViewerTheme } from 'pptx-react-viewer';

const custom: ViewerTheme = {
	colors: { ...vermilionDarkColors, primary: '#38bdf8' },
	radius: vermilionRadius,
};
```

Existe también una paleta clara sin marca (no vermilion), pero no es una exportación con nombre de los enlaces; se obtiene desde el catálogo de temas: `resolveThemeCatalogEntry('light')`.

## Aplicar un tema por enlace {#applying-a-theme-per-binding}

Todos los enlaces reciben el mismo objeto `ViewerTheme`; solo cambia el mecanismo de entrega.

::: code-group

```tsx [React]
import { PowerPointViewer, vermilionDarkTheme } from 'pptx-react-viewer';

<PowerPointViewer content={bytes} theme={vermilionDarkTheme} />;
```

```vue [Vue]
<script setup lang="ts">
import { PowerPointViewer, vermilionDarkTheme } from 'pptx-vue-viewer';
</script>

<template>
	<PowerPointViewer :content="bytes" :theme="vermilionDarkTheme" />
</template>
```

```ts [Angular]
// <pptx-viewer [content]="bytes" [theme]="theme" />
import { vermilionDarkTheme } from 'pptx-angular-viewer';

export class DeckComponent {
	theme = vermilionDarkTheme;
}
// Or share one theme across a subtree without the input:
// providers: [provideViewerTheme(vermilionDarkTheme)]
```

```svelte [Svelte]
<script lang="ts">
	import { PowerPointViewer, vermilionDarkTheme } from 'pptx-svelte-viewer';
</script>

<PowerPointViewer content={bytes} theme={vermilionDarkTheme} />
```

```ts [Vanilla]
import { createPptxViewer, vermilionDarkTheme } from 'pptx-vanilla-viewer';

const viewer = createPptxViewer(host, {
	source: bytes,
	theme: vermilionDarkTheme,
});
// Change later at runtime:
viewer.setTheme({ colors: { primary: '#38bdf8' } });
```

:::

Para compartir un mismo tema entre varios visualizadores, React exporta `ViewerThemeProvider` / `useViewerTheme`, Vue exporta `provideViewerTheme` / `useViewerTheme`, y Angular exporta `provideViewerTheme` / el token de inyección `VIEWER_THEME`.

## Modos de estilo {#styling-modes}

La interfaz del visualizador usa propiedades personalizadas `--pptx-*` para cada token visual, lo que permite tres formas de configurar los estilos (en orden creciente de control):

| Modo                                        | Configuración                                                                                                                                                                                                                                                         |
| ------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Host de Tailwind CSS v4                     | No hace falta importar CSS; las clases del visualizador se resuelven mediante su configuración existente. Sustituya los valores con la propiedad `theme`.                                                                                                             |
| Hoja de estilos empaquetada                 | `import 'pptx-react-viewer/styles'` (los paquetes de Vue y Angular también exponen `/styles` y `/styles.css`). El paquete de Svelte compila sus estilos dentro de sus componentes. Incluye todas las clases de utilidad necesarias y los valores oscuros por defecto. |
| Propiedades CSS personalizadas sin procesar | Defina usted mismo las propiedades `--pptx-*` (consulte `defaultCssVars()` para la lista completa) y omita las dos anteriores.                                                                                                                                        |

El enlace vanilla es la excepción: `createPptxViewer` inyecta automáticamente su propia hoja de estilos con ámbito (idempotente, `#pptx-vanilla-viewer-styles`). Los hosts con una CSP estricta pueden generar previamente la cadena a partir de `getViewerCss()`.

## Apariencia de los controles de selección {#selection-control-artwork}

La apariencia de la selección usa el campo existente `ViewerTheme.cssVars`, y no un objeto de tema ni una propiedad de componente aparte, en los cinco enlaces. Estos tokens opcionales no cambian la apariencia por defecto de un enlace cuando se omiten. Afectan a los controles del editor, no al contenido de las diapositivas, a los tamaños de fuente, a la geometría definida ni a los documentos exportados.

| Propiedad CSS personalizada            | Controla                                                            |
| -------------------------------------- | ------------------------------------------------------------------- |
| `--pptx-selection-corner-size`         | Ancho y alto de los marcadores de esquina                           |
| `--pptx-selection-corner-radius`       | Radio de borde de los marcadores de esquina                         |
| `--pptx-selection-edge-length`         | Longitud de los marcadores de borde, paralela al borde              |
| `--pptx-selection-edge-thickness`      | Grosor de los marcadores de borde, perpendicular al borde           |
| `--pptx-selection-edge-radius`         | Radio de borde de los marcadores de borde                           |
| `--pptx-selection-handle-fill`         | Relleno de los marcadores de cambio de tamaño                       |
| `--pptx-selection-handle-border-color` | Color del borde de los marcadores de cambio de tamaño y de rotación |
| `--pptx-selection-outline-color`       | Contorno de selección y vástago existente de rotación               |
| `--pptx-selection-rotate-size`         | Diámetro circular del marcador de rotación                          |
| `--pptx-selection-rotate-fill`         | Relleno del marcador de rotación                                    |
| `--pptx-selection-rotate-foreground`   | Color del glifo de rotación, cuando se dibuja un glifo              |

Use longitudes de píxeles positivas, como `6px`, para los tamaños; longitudes CSS o porcentajes para los radios; y colores CSS válidos para los colores. Una alternativa `var()` cubre un token omitido, no un valor CSS arbitrario no válido. El visualizador no analiza ni sanea estas cadenas CSS. Elimine una anulación para restaurar el valor predeterminado del enlace. Los valores de radio usan las coordenadas CSS locales del control, igual que su borde y su sombra; no se compensan por separado con el zoom del escenario. Use `0px` para esquinas cuadradas o `50%` para una forma redondeada proporcional en todos los enlaces.

```ts
const theme = {
	cssVars: {
		'--pptx-selection-corner-size': '6px',
		'--pptx-selection-corner-radius': '0px',
		'--pptx-selection-edge-length': '6px',
		'--pptx-selection-edge-thickness': '6px',
		'--pptx-selection-edge-radius': '0px',
		'--pptx-selection-handle-fill': '#ffffff',
		'--pptx-selection-handle-border-color': '#6366f1',
		'--pptx-selection-outline-color': '#6366f1',
		'--pptx-selection-rotate-fill': '#ffffff',
		'--pptx-selection-rotate-foreground': '#6366f1',
	},
};
```

Pase este mismo tema a través de los ejemplos de cada enlace que aparecen más arriba. Para un shell personalizado, aplique `themeToCssVars(theme)` al ancestro común de la diapositiva y de sus controles de selección. No dirija los estilos a clases hijas privadas ni suponga que los marcadores están anidados dentro de la forma seleccionada.

Los marcadores siguen centrados en los anclajes existentes. Un marcador más pequeño no reduce el área de ratón o de contacto invisible original. Un marcador más grande amplía su marco, pero la partición existente de la zona de los vecinos sigue aplicándose en formas muy pequeñas; los elementos visuales demasiado grandes pueden solaparse sin ocupar la zona de entrada de otro marcador. Este contrato no expone tamaños de zona de acción ni desplazamientos de rotación. Mantenga los marcadores personalizados grandes en un tamaño práctico para las formas que se editan. Los rombos de ajuste y los indicadores de extremo de conector conservan sus significados y apariencia propios.

## Archivo > Opciones > Apariencia {#file-options-appearance}

El cuadro de diálogo de Configuración de cada enlace tiene una pestaña **Apariencia**: una pequeña galería de predefinidos de tema integrados (Default, Light, Vermilion Light, Vermilion Dark) entre los que el usuario puede cambiar en tiempo de ejecución, definida por la exportación `THEME_CATALOG`:

```ts
interface ThemeCatalogEntry {
	key: string; // 'default' | 'light' | 'vermilionLight' | 'vermilionDark'
	labelKey: string; // pptx.* translation key for the entry's label
	theme: ViewerTheme | undefined; // undefined = reset to the built-in default
}
```

Es deliberadamente una lista corta y seleccionada, no una galería completa; pase `availableThemes` (más abajo) para ofrecer más o menos opciones. `resolveThemeCatalogEntry(key, catalog?)` busca una entrada por su clave.

### Prioridad: un `theme` explícito siempre gana {#precedence-an-explicit-theme-always-wins}

Si pasa una propiedad o entrada `theme` al visualizador, esta sigue prevaleciendo sobre cualquier opción elegida en la pestaña Apariencia; el selector no tiene efecto mientras `theme` esté definido. Es intencionado: un host que gestiona su propio tema (por ejemplo, sincronizado con un interruptor de modo oscuro de toda la aplicación) no debe verlo anulado por un clic dentro del visualizador.

La pestaña Apariencia solo surte efecto si **no** se pasa un `theme` explícito. En ese modo independiente, el orden de resolución es:

1. la propiedad `defaultThemeKey` (para un punto de partida distinto de `'default'`);
2. una elección guardada previamente (`localStorage`, clave `pptx-viewer-prefs`);
3. la entrada `'default'` del catálogo.

### Propiedades del catálogo (todas opcionales) {#catalog-props-all-optional}

| Propiedad          | Tipo                     | Finalidad                                                                                                                                                                                                                                                                                                      |
| ------------------ | ------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `defaultThemeKey`  | `string`                 | Clave inicial de `THEME_CATALOG`, que se usa solo cuando todavía no hay nada guardado.                                                                                                                                                                                                                         |
| `availableThemes`  | `ThemeCatalogEntry[]`    | Sustituye el catálogo que se ofrece en Apariencia (añada sus propios predefinidos o reduzca la lista).                                                                                                                                                                                                         |
| `onThemeChange`    | `(key: string) => void`  | Si se proporciona, **usted** se encarga de guardar la elección; el visualizador deja de escribir en `localStorage` y solo le llama de vuelta.                                                                                                                                                                  |
| `defaultLocale`    | `string`                 | La misma idea para el idioma: código de configuración regional inicial cuando no hay nada guardado. React: se ignora dentro de un `PptxI18nProvider`, cuyo `locale` es el idioma.                                                                                                                              |
| `availableLocales` | `LocaleCatalogEntry[]`   | Sustituye las configuraciones regionales que se ofrecen en Opciones > Idioma. React: por defecto, los códigos registrados en la instancia de `react-i18next` del host; dentro de un `PptxI18nProvider` no hay valor por defecto: pase dos o más, junto con `onLocaleChange`, para mostrar la categoría Idioma. |
| `onLocaleChange`   | `(code: string) => void` | Si se proporciona, el visualizador nunca toca su propia instancia de i18n; solo le llama de vuelta. React: dentro de un `PptxI18nProvider` el host siempre gestiona el idioma, y sin este callback la categoría Idioma se omite.                                                                               |

React, Vue, Angular y Svelte usan exactamente esta misma forma. Vanilla es la excepción: ya tenía las opciones públicas `theme`/`locale` del constructor y los métodos `setTheme()`/`setLocale()`, así que estos sirven como valor inicial en lugar de las opciones separadas `defaultThemeKey`/`defaultLocale`; `availableThemes`, `availableLocales`, `onThemeChange` y `onLocaleChange` son iguales que en el resto.

```tsx
// React - host owns theme persistence (e.g. syncing with app-wide dark mode)
<PowerPointViewer
	content={bytes}
	defaultThemeKey={systemPrefersDark ? 'vermilionDark' : 'vermilionLight'}
	onThemeChange={(key) => saveUserPreference('viewerTheme', key)}
/>
```

### Añadir sus propios predefinidos {#adding-your-own-presets}

Pase `availableThemes` con su propio `ThemeCatalogEntry[]`; amplíe los predefinidos integrados en lugar de sustituirlos, desplegando `THEME_CATALOG`:

```ts
import { THEME_CATALOG } from 'pptx-react-viewer';

const availableThemes = [
	...THEME_CATALOG,
	{ key: 'brand', labelKey: 'app.theme.brand', theme: { colors: { primary: '#7c3aed' } } },
];
```

`labelKey` se busca a través del diccionario de i18n que proporcione su aplicación (consulte [Localización](/es/guide/localization)); para una entrada personalizada, registre esa clave usted mismo. Las claves no registradas recurren a una etiqueta legible derivada del último segmento de la clave.

## Detalles por enlace {#per-binding-details}

- [Temas de React](/react/theming) - la propiedad `theme`, `ViewerThemeProvider` y la tabla completa de tokens
- [Temas de Vue](/vue/theming) - la propiedad `theme` y `provideViewerTheme`
- [Temas de Angular](/angular/theming) - la entrada `theme`, `provideViewerTheme` y el token `VIEWER_THEME`
- [Temas de Svelte](/svelte/theming) - la propiedad `theme`
- [Temas de Vanilla](/vanilla/theming) - la opción `theme`, `setTheme()` y `getViewerCss()`

## Siguientes pasos {#next-steps}

- [Localización (i18n)](/es/guide/localization) - la pestaña Idioma funciona igual, y de ahí proceden las claves de traducción `pptx.*` que respaldan las etiquetas de los temas.
- [Cuenta e inicio de sesión](/es/guide/account) - el editor de perfil de File > Account también guarda en la misma clave de `localStorage` que la alternativa de tema e idioma.
