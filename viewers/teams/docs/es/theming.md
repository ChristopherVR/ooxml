# Temas

`<teams-app>` toma cada color, valor de espaciado y radio de los tokens compartidos de `ooxml-ui`
(`var(--office-...)`). Las propiedades personalizadas atraviesan los límites de shadow DOM, así que
una sola regla en `:root` (o en un contenedor) aplica el tema a toda la aplicación.

## Claro y oscuro {#light-and-dark}

Los tokens siguen por defecto el esquema de color del sistema operativo. Para forzar un esquema,
establece `data-office-theme` en la raíz:

```ts
document.documentElement.setAttribute('data-office-theme', 'dark'); // or 'light'
document.documentElement.removeAttribute('data-office-theme'); // follow the OS again
```

## Sobrescritura de tokens {#token-overrides}

Sobrescribe cualquier token donde quieras que se aplique:

```css
:root {
	--office-accent: #0e8f8f;
	--office-teams-brand: #0e8f8f;
	--office-teams-rail: #14171b;
}
```

Los tokens específicos de Teams llevan el prefijo `--office-teams-` (marca, barra lateral, panel,
burbuja, divisor y texto secundario). La aplicación no tiene una propiedad `theme` propia: el tema se
define con estos tokens. El color de énfasis de la documentación de este sitio es un verde azulado
neutro a propósito; OpenTeams no toma prestados los colores de marca de Microsoft.

## Seguir el tema del sitio de documentación {#following-the-docs-site-theme}

Este sitio guarda la preferencia de apariencia del visitante en `localStorage` con la clave de
VitePress `vitepress-theme-appearance` (`light`, `dark` o `auto`). Todos los sitios de visores
comparten el origen `christophervr.github.io`, y las aplicaciones de demo de este repositorio
(`demos/teams/theme-sync.ts` en la raíz del repositorio) la aplican al arrancar y al recibir el evento
`storage`, de modo que cambiar el tema de este sitio vuelve a dar estilo a una demo incrustada sin
recargar:

```ts
const key = 'vitepress-theme-appearance';
const apply = (value: string | null) => {
	const root = document.documentElement;
	if (value === 'light' || value === 'dark') root.setAttribute('data-office-theme', value);
	else root.removeAttribute('data-office-theme');
};
apply(localStorage.getItem(key));
window.addEventListener('storage', (event) => {
	if (event.key === key || event.key === null) apply(event.newValue);
});
```
