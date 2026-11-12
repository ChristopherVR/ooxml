# Las demos en vivo

La misma aplicación de demostración se compila una vez por cada adaptador de framework y se publica junto a estos documentos. Cada una es el `<docx-editor>` real ejecutándose por completo en su navegador: no se sube nada y no hay ningún servidor detrás.

| Demo                                                | Adaptador             | Origen                                                              |
| --------------------------------------------------- | --------------------- | ------------------------------------------------------------------- |
| [React](/demo/){target="_self"}                     | `docx-react-viewer`   | `demos/docx/demo-vanilla` compilada con `VITE_DEMO_FRAMEWORK=react` |
| [Vue](/demo-vue/){target="_self"}                   | `docx-vue-viewer`     | la misma aplicación, `vue`                                          |
| [Angular](/demo-angular/){target="_self"}           | `docx-angular-viewer` | la misma aplicación, `angular`                                      |
| [Svelte](/demo-svelte/){target="_self"}             | `docx-svelte-viewer`  | la misma aplicación, `svelte`                                       |
| [Solid](/demo-solid/){target="_self"}               | `docx-solid-viewer`   | la misma aplicación, `solid`                                        |
| [JavaScript nativo](/demo-vanilla/){target="_self"} | `docx-vanilla-viewer` | la misma aplicación, `vanilla`                                      |

La página de inicio las incrusta con un selector de framework: [abra la sección de demo en vivo](/es/#live-demo).

## Qué puede hacer en una demo {#what-you-can-do-in-a-demo}

- Elija **Abrir el documento de ejemplo** o arrastre un archivo DOCX o DOC heredado propio. Los archivos permanecen en la pestaña.
- Añada `?sample=1` a la URL de una demo para abrir el documento de ejemplo directamente, y `?locale=de` (también `fr`, `es`, `zh-CN`) para iniciar la interfaz en otro idioma.
- Use el botón de tema de la demo o el selector de tema de este sitio: todas las demos siguen en directo la preferencia `vitepress-theme-appearance` (consulte [temas](/es/theming)).

## Demo de colaboración {#collaboration-demo}

Añada `?room=<session>` a la URL de una demo para unirse a una sesión compartida del mismo navegador. La ventana que además tiene `sample=1` aloja la sesión; las demás se unen por nombre, en cualquier framework: por ejemplo, `/demo-vue/?sample=1&room=my-room` y `/demo-solid/?room=my-room` en dos pestañas. La página de inicio muestra dos paneles con un selector de framework para cada uno. La [página de dos pares](/demo/collaboration.html){target="_self"} añade un botón de Pausar la entrega para las ediciones simultáneas. Consulte [colaboración](/es/collaboration) para ver qué se admite y qué no.

## Incrustada aquí {#embedded-here}

<iframe
	src="/ooxml/docx/demo/?sample=1"
	title="docx-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Ejecútelas en local {#run-them-locally}

```bash
bun install          # from the repository root
cd viewers/docx
bun run demo
```

`bun run docs:build` (desde `docs/`) compila la documentación y las seis demos exactamente como lo hace el flujo de trabajo de Pages.
