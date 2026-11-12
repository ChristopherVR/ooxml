# Las demos en vivo

La misma aplicación de demostración se compila una vez por cada adaptador de framework y se publica junto a estas páginas. Cada una es el `<xlsx-editor>` real que se ejecuta por completo en su navegador: los libros se analizan, editan y guardan en la pestaña, nada se sube y no hay ningún servidor detrás.

| Demo                                              | Adaptador                          |
| ------------------------------------------------- | ---------------------------------- |
| [React](/demo/){target="_self"}                   | `@christophervr/xlsx-react-viewer` |
| [Vue](/demo-vue/){target="_self"}                 | `xlsx-vue-viewer`                  |
| [Angular](/demo-angular/){target="_self"}         | `xlsx-angular-viewer`              |
| [Svelte](/demo-svelte/){target="_self"}           | `xlsx-svelte-viewer`               |
| [Solid](/demo-solid/){target="_self"}             | `xlsx-solid-viewer`                |
| [JavaScript puro](/demo-vanilla/){target="_self"} | `xlsx-vanilla-viewer`              |

La página de inicio las incrusta con un selector de framework: [abra la sección de demo en vivo](/es/#live-demo).

## Qué puede hacer en una demo {#what-you-can-do-in-a-demo}

- Suelte un archivo `.xlsx`, `.xlsm`, `.xls` heredado o `.csv`, cree un libro nuevo o abra el libro de ejemplo.
- Añada `?sample=1` para abrir de inmediato el libro de ejemplo, y `?locale=de` (cualquier etiqueta admitida) para empezar con otro idioma de la interfaz.
- Use el botón de tema de la demo o el selector de tema de este sitio: las demos siguen en vivo la preferencia `vitepress-theme-appearance` (vea [temas](/es/theming)).

## Colaboración {#collaboration}

La demo no tiene una página de colaboración aparte: la coedición está integrada en el editor. Elija **Share** (Compartir) y después inicie una sesión; otra ventana o pestaña del mismo navegador se une desde el enlace de invitación o por el nombre de la sesión. Vea [colaboración](/es/collaboration) para las opciones de sala y lo que se comparte y lo que no.

## Incrustada aquí {#embedded-here}

<iframe
	src="/ooxml/xlsx/demo/?sample=1"
	title="xlsx-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Ejecútelas en local {#run-them-locally}

```bash
bun install   # at the repository root
cd viewers/xlsx
bun run demo
```

`bun run docs:build` (desde `docs/`) compila la documentación y las seis demos exactamente igual que el flujo de Pages.
