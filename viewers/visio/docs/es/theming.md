# Temas

El visor dibuja su propia interfaz (cinta, backstage, barra de estado, ventana de Formas) a partir de controles compartidos de `ooxml-ui`. El elemento lee sus colores de propiedades personalizadas de CSS, de modo que un anfitrión puede cambiar su estilo desde su propia hoja de estilos.

## Tokens `--office-*` compartidos {#shared-office-tokens}

El elemento sigue el tema Office compartido. Al conectarse instala el tema de `ooxml-ui` (`installOfficeUiTheme`, idempotente), y cada color que dibuja se resuelve a partir de las propiedades personalizadas `--office-*`, por lo que los tokens `--office-*` son la forma principal de darle tema. Defínalos en `:root`, en un contenedor o en el propio elemento:

| Token                        | Función                                            |
| ---------------------------- | -------------------------------------------------- |
| `--office-background`        | Superficie de la cinta, los paneles y los diálogos |
| `--office-surface`           | Fondo de la ventana, superficie secundaria         |
| `--office-foreground`        | Texto principal                                    |
| `--office-muted-foreground`  | Texto secundario                                   |
| `--office-border`            | Bordes                                             |
| `--office-accent`            | Color de énfasis                                   |
| `--office-accent-foreground` | Texto sobre el color de énfasis                    |
| `--office-ring`              | Anillo de foco                                     |
| `--office-danger`            | Estados destructivos y de error                    |
| `--office-shadow`            | Sombra de la página                                |

```css
visio-viewer {
	--office-accent: #c2431f;
	--office-accent-foreground: #ffffff;
}
```

El modo claro y el oscuro siguen el tema de la suite: el visor es claro por defecto, oscuro cuando el sistema lo prefiere, y el atributo `data-office-theme="light" | "dark"` en la raíz del documento fuerza uno de los dos. El elemento no tiene una propiedad `theme`; el mismo selector que da tema a los demás elementos de Office de la página da tema también al visor.

El color de énfasis por defecto ha cambiado con esto. Antes el visor tenía fijo `#3955a3`; ahora toma `--office-accent` (`#2563eb` en claro y `#6366f1` en oscuro). Un anfitrión que quiera el color anterior define `--office-accent: #3955a3` o la sobrescritura `--vv-accent` que aparece más abajo.

## Sobrescrituras opcionales `--vv-*` {#optional-vv-overrides}

Los tokens `--vv-*` anteriores siguen funcionando. Cada uno, cuando se define, prevalece sobre el token `--office-*` al que equivale; cuando no se define, el visor recurre al token compartido. Úselos para dar estilo solo a este visor sin tocar el resto de la página.

| Token              | Recurre a                    | Función                                                         |
| ------------------ | ---------------------------- | --------------------------------------------------------------- |
| `--vv-font`        | `--office-font`              | Familia tipográfica                                             |
| `--vv-background`  | `--office-surface`           | Fondo de la ventana                                             |
| `--vv-surface`     | `--office-background`        | Superficie de la cinta y los paneles                            |
| `--vv-secondary`   | `--office-surface`           | Superficie secundaria                                           |
| `--vv-shadow`      | `--office-shadow`            | Sombra de la página                                             |
| `--vv-ink`         | `--office-foreground`        | Texto principal                                                 |
| `--vv-muted`       | `--office-muted-foreground`  | Texto secundario                                                |
| `--vv-border`      | `--office-border`            | Bordes                                                          |
| `--vv-accent`      | `--office-accent`            | Color de énfasis                                                |
| `--vv-accent-soft` | 10 % del énfasis             | Superficie de énfasis suave (al pasar el puntero, seleccionado) |
| `--vv-accent-ink`  | `--office-accent-foreground` | Texto sobre el color de énfasis                                 |
| `--vv-focus`       | `--office-ring`              | Anillo de foco                                                  |
| `--vv-danger`      | `--office-danger`            | Estados destructivos y de error                                 |

El área de pruebas define `data-theme="light" | "dark"` en la raíz del documento y aporta valores `--vv-*` coincidentes en `demos/visio/demo/workspace.css` para la interfaz de su propia página (pantalla de inicio, shell del espacio de trabajo); ese archivo sigue funcionando porque las sobrescrituras `--vv-*` siguen aplicándose.

## Seguir el tema del sitio de documentación {#following-the-docs-site-theme}

Este sitio guarda la elección de apariencia del visitante en `localStorage` con la clave de VitePress `vitepress-theme-appearance` (`light`, `dark` o ningún valor, que sigue al sistema). Todos los sitios de los visores comparten el origen `christophervr.github.io`, y el área de pruebas (`demos/visio/demo/suite-theme.js`) lee la misma clave al arrancar y se actualiza en tiempo real con el evento `storage`, de modo que cambiar el tema en este sitio vuelve a dar estilo a un visor incrustado sin recargarlo ni sustituir su diagrama.

```ts
const key = 'vitepress-theme-appearance';
const resolve = (value: string | null) =>
	value === 'dark' || value === 'light'
		? value
		: matchMedia('(prefers-color-scheme: dark)').matches
			? 'dark'
			: 'light';

document.documentElement.dataset.theme = resolve(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null)
		document.documentElement.dataset.theme = resolve(event.newValue);
});
```

## Tema de la barra de estado {#status-bar-theme}

La barra de estado es neutra por defecto, como los demás editores de la suite. El tema Office «Colorido» de Visio de escritorio la rellena con el color del producto. Hay tres formas de elegir:

- **El usuario**: Archivo > Opciones > General > Tema de Office (Colorido, Blanco o Negro, que es el tema oscuro). La elección se guarda en el dispositivo.
- **El anfitrión**: el atributo `status-bar` (`colorful` o `neutral`), o la propiedad `statusBar`, en `<visio-viewer>`. El atributo de un anfitrión prevalece sobre la elección guardada.
- **Cualquier color**: `--vv-status-background` y `--vv-status-ink` cambian el color de la barra en cualquiera de los dos temas.

```html
<visio-viewer status-bar="colorful"></visio-viewer>
<visio-viewer style="--vv-status-background: #1b5e20; --vv-status-ink: #fff"></visio-viewer>
```
