# Las demostraciones en vivo

El mismo área de pruebas se compila una vez por cada adaptador de framework y se publica junto a estos documentos. Cada una es el `<visio-viewer>` real que se ejecuta por completo en su navegador: los dibujos se analizan y se representan de forma local, no se sube nada y no hay ningún servidor detrás. Esta es una beta pública; que una demostración abra un archivo no es prueba de la fidelidad con Visio.

| Demostración                                 | Adaptador              | Ruta             |
| -------------------------------------------- | ---------------------- | ---------------- |
| [JavaScript vanilla](/demo/){target="_self"} | `visio-vanilla-viewer` | `/demo/`         |
| [React](/demo-react/){target="_self"}        | `visio-react-viewer`   | `/demo-react/`   |
| [Vue](/demo-vue/){target="_self"}            | `visio-vue-viewer`     | `/demo-vue/`     |
| [Angular](/demo-angular/){target="_self"}    | `visio-angular-viewer` | `/demo-angular/` |
| [Svelte](/demo-svelte/){target="_self"}      | `visio-svelte-viewer`  | `/demo-svelte/`  |
| [Solid](/demo-solid/){target="_self"}        | `visio-solid-viewer`   | `/demo-solid/`   |

El área de pruebas vanilla también se sirve en [/demo-vanilla/](/demo-vanilla/){target="_self"}, el nombre que usan los sitios de los demás visores. La página de inicio incrusta las demostraciones con un selector de framework: [abra la sección de demostración en vivo](/es/#live-demo).

## Qué puede hacer en una demostración {#what-you-can-do-in-a-demo}

- Abrir el dibujo de ejemplo original, o un archivo `.vsdx` local que elija.
- Añadir `?sample=1` a la URL de una demostración para abrir el ejemplo de inmediato, y `?embed=1` para la disposición compacta incrustada que usa la página de inicio.
- Usar el botón de tema de la demostración o el selector de tema de este sitio: todas las demostraciones siguen la preferencia `vitepress-theme-appearance` en tiempo real (consulte [temas](/es/theming)).

## Demostración de colaboración {#collaboration-demo}

Añada `?share=<session>` a cualquier URL de demostración para unirse a esa sesión de uso compartido al cargar. Dos demostraciones creadas con frameworks distintos que usan el mismo nombre de sesión comparten un dibujo; por ejemplo, `/demo-vue/?sample=1&share=my-room` y `/demo-angular/?share=my-room` en dos pestañas de un mismo navegador. La página de inicio muestra dos paneles con un selector de framework para cada uno. Consulte [colaboración](/es/collaboration) para ver qué se admite y qué no.

## Incrustada aquí {#embedded-here}

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1"
	title="visio-viewer live demo"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

## Ejecútelas en local {#run-them-locally}

```bash
bun install            # once, at the repository root
cd viewers/visio
bun run dev
```

El código fuente de las demostraciones está en `demos/visio/` en la raíz del repositorio (`demo` es el área de pruebas vanilla compartida; `demo-react`, `demo-vue`, `demo-angular`, `demo-svelte` y `demo-solid` montan el mismo espacio de trabajo). Los scripts del visor y su configuración de Vite las compilan y las sirven. `bun run docs:build` compila la documentación y todas las demostraciones exactamente como lo hace la compilación de Pages del repositorio (`node scripts/build-pages.mjs` en la raíz compila el sitio de cada visor).
