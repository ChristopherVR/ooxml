# Publicar los paquetes

Los paquetes de Excel se publican mediante el flujo automatizado único del repositorio, no mediante un flujo propio. Este visor es `viewers/xlsx` del repositorio [ooxml](https://github.com/ChristopherVR/ooxml), junto con `ooxml-core` (`src/core`) y `ooxml-ui` (`src/ui`). Todo el proceso (cómo se decide una publicación, el nivel de incremento de versión, el flujo horario, la publicación confiable y la configuración única del repositorio) se describe en [docs/releasing.md](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md) en la raíz del repositorio.

## Qué se publica desde aquí {#what-is-published-from-here}

Cada paquete tiene su propia línea de versiones, su etiqueta git `<npm-name>@<version>` y su changelog. Se declaran en `scripts/viewer-packages.mjs` en la raíz del repositorio.

| Paquete                            | Directorio         | Notas                                                                            |
| ---------------------------------- | ------------------ | -------------------------------------------------------------------------------- |
| `@christophervr/xlsx-core`         | `packages/core`    | Punto de entrada del modelo de libro sobre `ooxml-core`; no tiene lógica propia. |
| `@christophervr/xlsx-react-viewer` | `packages/react`   | Paquete de editor autocontenido; incluye el componente web y los bindings.       |
| `xlsx-vue-viewer`                  | `packages/vue`     | Misma forma.                                                                     |
| `xlsx-angular-viewer`              | `packages/angular` | Misma forma.                                                                     |
| `xlsx-svelte-viewer`               | `packages/svelte`  | Misma forma.                                                                     |
| `xlsx-solid-viewer`                | `packages/solid`   | Misma forma.                                                                     |
| `xlsx-vanilla-viewer`              | `packages/vanilla` | Misma forma; también es dueño de la entrada `<xlsx-editor>` simple.              |
| `xlsx-viewer-mcp`                  | `mcp`              | Servidor MCP independiente, vea su README.                                       |

Los paquetes internos `packages/web-component` (`xlsx-web-component`) y `packages/bindings` (`xlsx-bindings`) son `private` y nunca se publican: `scripts/build-packages.mjs` los incluye en cada paquete de framework, así que un cambio que no sea de prueba en cualquiera de los dos publica los seis. Un tarball publicado solo importa `@christophervr/xlsx-core`, `ooxml-core`, `ooxml-ui` y su peer del framework; `ole2` va incluido dentro de `ooxml-core`. `ooxml-core` y `ooxml-ui` son dependencias reales del registro, nunca empaquetadas.

## Comprobaciones desde esta carpeta {#checks-from-this-folder}

Ejecútelas desde `viewers/xlsx` tras un `bun install` en la raíz:

```sh
bun run build:packages
bun run check:published   # no tarball imports an internal package or ole2
bun run pack:smoke        # install the tarballs into a clean consumer and load .xlsx and .xls
```

El planificador de publicaciones, las guardas de publicación y el flujo se ejecutan desde la raíz del repositorio (`bun run release:plan`, `node scripts/publish-released.mjs --dry-run`, `.github/workflows/release.yml`). Nunca ejecute `npm publish` ni suba etiquetas de publicación a mano.

## Convenciones de commit {#commit-conventions}

El nivel de incremento de versión lo determina el tipo del commit (`!` o `BREAKING CHANGE:` es mayor, `feat` es menor, cualquier otro es parche), y los archivos que toca un commit deciden si un paquete se publica. Las reglas y los ejemplos están en [CONTRIBUTING.md](https://github.com/ChristopherVR/ooxml/blob/main/CONTRIBUTING.md#commit-conventions).
