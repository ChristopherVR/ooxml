# Temas

El editor da tema a su propia interfaz (barra de título, cinta, paneles, barra de estado, diálogos). El papel del
documento siempre permanece blanco, como la página de Word.

## Claro, oscuro y automático {#light-dark-and-auto}

Defina la propiedad `theme` (atributo `theme`) como `light`, `dark` o `auto`. Los valores desconocidos recurren a
`auto`, que hereda el ámbito compartido de Office más cercano y, si no hay ninguno, sigue el esquema de colores del
sistema operativo.

```ts
editor.theme = 'dark';
```

## Tokens compartidos de Office {#shared-office-tokens}

Word usa los mismos tokens `--office-*` que Excel, Visio, PowerPoint y los controles compartidos.
Defina `data-office-theme="light"`, `"dark"` o `"auto"` en un ancestro común para dar tema a una composición
completa. Su modo y sus tokens personalizados se actualizan en directo, también a través del shadow root del editor.
Los componentes nativos de barra de herramientas de React, Vue, Angular, Svelte y Solid adoptan los mismos estilos compartidos.

```html
<section data-office-theme="dark" class="office-host">
	<!-- native canvas and toolbar components -->
</section>
<style>
	.office-host {
		--office-accent: #9b59b6;
		--office-ring: #b37cd0;
		--office-radius: 6px;
		--office-font: system-ui, sans-serif;
	}
</style>
```

Use `--office-background` para el marco de la ventana, `--office-surface` para los paneles y el área de trabajo,
`--office-foreground` para el texto de la interfaz, y `--office-border` para los contornos. Los controles integrados
también comparten los tokens de tipografía, espaciado, foco y estado deshabilitado de Office. Los anfitriones pueden
sustituir los valores predeterminados de baja especificidad con sus propios estilos. El formato de texto del documento
y el tema de Office serializado no cambian. El papel sigue siendo blanco por defecto en modo oscuro; `--office-paper`
y el `--dve-page-color` existente ofrecen sustituciones explícitas de presentación.

Un `theme="light"` o `theme="dark"` explícito en el editor crea un ámbito local para su lienzo.
Use el ancestro común para dar tema también a los controles nativos hermanos; un editor no cambia el tema de su
aplicación anfitriona. Devolver el editor a `auto` restablece el comportamiento de ámbito heredado.

## Sustituciones de tokens heredados {#legacy-token-overrides}

Las sustituciones `--dve-*` siguen siendo compatibles y tienen prioridad sobre su valor de reserva canónico.
`themeColors` acepta un `EditorTheme` parcial y aplica cada clave como propiedad personalizada `--dve-*` en el
anfitrión, tanto en claro como en oscuro. Los valores pueden ser cualquier color CSS.

```ts
editor.themeColors = {
	primary: '#c2431f',
	primaryForeground: '#ffffff',
	radius: '6px',
};
```

| Token                                  | Función                                                    |
| -------------------------------------- | ---------------------------------------------------------- |
| `background`, `foreground`             | Interfaz de la aplicación (barra de título, ventana)       |
| `card`, `cardForeground`               | Superficie de la cinta, los paneles y la barra de estado   |
| `popover`, `popoverForeground`         | Menús, diálogos y otras superficies flotantes              |
| `primary`, `primaryForeground`         | Color de marca y de la acción principal                    |
| `secondary`, `secondaryForeground`     | Superficies secundarias                                    |
| `muted`, `mutedForeground`             | Superficie atenuada, también el área detrás de las páginas |
| `accent`, `accentForeground`           | Resaltado al pasar el cursor y de selección                |
| `destructive`, `destructiveForeground` | Acciones destructivas                                      |
| `border`, `input`, `ring`              | Bordes, contornos de campos, anillo de foco                |
| `radius`                               | Radio base de las esquinas, por ejemplo `4px`              |

Pase `undefined` para borrar las sustituciones.

## Seguir el tema del sitio de documentación {#following-the-docs-site-theme}

Este sitio guarda la preferencia de apariencia del visitante en `localStorage`, con la clave de VitePress
`vitepress-theme-appearance` (`light`, `dark` o `auto`). Todos los sitios del visor comparten el origen
`christophervr.github.io`, y las aplicaciones de demostración leen la misma clave al arrancar y se actualizan en
directo con el evento `storage`, de modo que cambiar el tema en este sitio vuelve a dar estilo a un editor
incrustado sin recargarlo. En resumen:

```ts
const key = 'vitepress-theme-appearance';
const resolve = (value: string | null) =>
	value === 'dark' || value === 'light'
		? value
		: matchMedia('(prefers-color-scheme: dark)').matches
			? 'dark'
			: 'light';

editor.theme = resolve(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null) editor.theme = resolve(event.newValue);
});
```

Una aplicación anfitriona puede usar su propia clave: el elemento solo necesita la propiedad `theme`.
