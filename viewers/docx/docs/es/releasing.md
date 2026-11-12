# Publicación de paquetes

Los paquetes del visor de Word los publica el propio flujo del repositorio `ooxml`, junto con `ooxml-core`
y `ooxml-ui`. No hay un flujo de publicación separado en `viewers/docx`. Todo el flujo (el planificador, el
`release.yml` horario, la publicación de confianza, los registros de cambios y la configuración inicial única)
se describe en la [guía de publicación](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md)
del repositorio. Nunca ejecute `npm publish` ni envíe etiquetas de publicación a mano.

## Lo que publica este visor {#what-this-viewer-publishes}

Cada paquete tiene su propia línea de versiones y su propia etiqueta de git `<npm-name>@<version>`. La
correspondencia entre directorio y nombre de npm está en `scripts/viewer-packages.mjs`, en la raíz del repositorio.

| Paquete               | Directorio         | Notas                                                            |
| --------------------- | ------------------ | ---------------------------------------------------------------- |
| `docx-core`           | `packages/core`    | Reexportación fina de `ooxml-core/docx` y `/docx/embedded`.      |
| `docx-react-viewer`   | `packages/react`   | Entrada de React (`WordEditor`).                                 |
| `docx-vue-viewer`     | `packages/vue`     | Entrada de Vue 3.                                                |
| `docx-angular-viewer` | `packages/angular` | Entrada de Angular (`WordEditorComponent`).                      |
| `docx-svelte-viewer`  | `packages/svelte`  | Entrada de Svelte 5.                                             |
| `docx-solid-viewer`   | `packages/solid`   | Entrada de Solid.                                                |
| `docx-vanilla-viewer` | `packages/vanilla` | Registro de `<docx-editor>` nativo y `mountEditor`.              |
| `docx-viewer-mcp`     | `mcp`              | El servidor MCP de Word (consulte su README y el `mcp.md` raíz). |

`web-component` y `bindings` son `private: true` y nunca se publican.
`scripts/build-packages.mjs` los empaqueta en cada paquete de framework, de modo que un tarball solo importa
`docx-core`, `ooxml-core`, `ooxml-ui`, las bibliotecas de ProseMirror y su peer de framework.
`ooxml-ui` y `docx-core` siguen siendo dependencias reales del registro (nunca empaquetadas), y los códecs
heredados `.doc` de `ole2` van incrustados dentro de `ooxml-core`, de modo que `ole2` no es una dependencia de nada de aquí.

## Qué desencadena una publicación {#what-triggers-a-release}

El planificador raíz (`scripts/release-plan.mjs`, usando la tabla de `scripts/viewer-packages.mjs`) revisa los
commits convencionales desde la última etiqueta de cada paquete:

- Un cambio que no sea de pruebas dentro de `packages/<name>/` publica ese paquete.
- Un cambio que no sea de pruebas dentro de `packages/web-component` o `packages/bindings` publica todos los
  paquetes de framework, porque cada uno incluye ese código. `scripts/build-packages.mjs` y `tsconfig.release.json`
  publican todo lo de este visor.
- Una publicación de `docx-core` vuelve a publicar los paquetes de framework que dependen de él.
- Las pruebas, la documentación, las demos y la CI nunca publican nada. Un cambio en `demos/docx/` o `e2e/docx/`
  solo ejecuta las comprobaciones de este visor en la CI.

El código del propio editor vive en `src/ui/src/docx` y se distribuye en `ooxml-ui`; un cambio allí publica
`ooxml-ui`, y los paquetes del visor lo recogen a través de su rango de `ooxml-ui`.

## Comprobaciones antes de enviar {#checks-before-pushing}

Ejecute `bun run verify` en la raíz del repositorio. Para este visor, las comprobaciones locales equivalentes,
desde `viewers/docx`, son:

```sh
bun run typecheck && bun run test && bun run test:scripts
bun run build:packages && bun run check:published && bun run pack:smoke
```

`check:published` falla si un tarball importa un paquete interno o `ole2`, y `pack:smoke` instala los tarballs en
un consumidor limpio y carga DOCX y `.doc` heredado a través de cada paquete de framework.
