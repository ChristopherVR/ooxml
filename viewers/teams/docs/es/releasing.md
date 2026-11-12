# Publicar los paquetes

Siete paquetes sin ámbito se publican en npm con el nombre del producto **OpenTeams** (`openteams`
en sí pertenece a un proyecto sin relación, por eso cada nombre es más largo). Cada uno tiene su
**propia línea de versiones**, que solo se incrementa cuando cambia de verdad, y su propia etiqueta
de git `<nombre-npm>@<versión>` (por ejemplo `openteams-server@0.4.0`).

Las publicaciones se ejecutan desde la **raíz del repositorio**, no desde esta carpeta. El
planificador (`scripts/release-plan.mjs`), la tabla de paquetes del visor (`scripts/viewer-packages.mjs`),
el flujo de trabajo (`.github/workflows/release.yml`), la publicación de confianza y la configuración
inicial se describen en la [`docs/releasing.md`](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md)
de la raíz. Esta carpeta no tiene scripts de publicación, ni flujo de trabajo propio (su `.github/`
está inactivo) ni los comandos `release:plan` o `publish:local`.

| Paquete                    | Directorio         | Depende de (registro)                                 | Incluido en                     |
| -------------------------- | ------------------ | ----------------------------------------------------- | ------------------------------- |
| `openteams-react-viewer`   | `packages/react`   | `ooxml-core`, `ooxml-ui`, `lit`; peer `react`         | `packages/web-component`        |
| `openteams-vue-viewer`     | `packages/vue`     | `ooxml-core`, `ooxml-ui`, `lit`; peer `vue`           | `packages/web-component`        |
| `openteams-angular-viewer` | `packages/angular` | `ooxml-core`, `ooxml-ui`, `lit`; peer `@angular/core` | `packages/web-component`        |
| `openteams-svelte-viewer`  | `packages/svelte`  | `ooxml-core`, `ooxml-ui`, `lit`; peer `svelte`        | `packages/web-component`        |
| `openteams-solid-viewer`   | `packages/solid`   | `ooxml-core`, `ooxml-ui`, `lit`; peer `solid-js`      | `packages/web-component`        |
| `openteams-vanilla-viewer` | `packages/vanilla` | `ooxml-core`, `ooxml-ui`, `lit`                       | `packages/web-component`        |
| `openteams-server`         | `server`           | `ws`, `yjs`, `y-protocols`, `lib0`                    | nada (publica su código fuente) |

`teams-viewer` (`packages/web-component`, una reexportación de `ooxml-ui/teams`, que contiene
`<teams-app>` y el cableado de los bindings) y las demos son `private` y nunca se publican.
`scripts/build-packages.mjs` incluye el componente web, con su CSS, en cada paquete de framework
(compilación de biblioteca con vite) y aplana sus declaraciones en el `dist/index.d.ts` de cada
paquete (rollup-plugin-dts), de modo que un tarball importa solo lo que declara su manifiesto.
Svelte publica `dist/Teams.svelte` como código fuente junto a un `dist/runtime.js` empaquetado (el
store y las utilidades, bajo `openteams-svelte-viewer/runtime`); Angular es un módulo ES empaquetado
con sus decoradores aplicados en tiempo de ejecución (necesita el compilador JIT en la aplicación que
lo consume). El servidor es ESM simple y no se compila; su comando `openteams-server` es
`server/bin.mjs`.

En el monorepositorio, los manifiestos de los frameworks dependen de `ooxml-core` y `ooxml-ui` con
`*`; el flujo de publicación de la raíz reescribe esos rangos como rangos con caret sobre las
versiones publicadas junto con ellos. Un cambio que necesite a la vez el núcleo y este visor se
integra en un único envío y se publica conjuntamente.

## Qué decide una publicación aquí {#what-decides-a-release-here}

El planificador de la raíz lee `bundled` y `globals` de la entrada `teams` de
`scripts/viewer-packages.mjs`:

- **Archivos propios.** Cualquier archivo que no sea de prueba bajo el directorio de un paquete lo
  publica. Las pruebas, `CHANGELOG.md` y las ediciones del manifiesto que solo tocan `version`,
  `scripts` o `devDependencies` nunca lo hacen.
- **Paquete interno incluido.** Un cambio que no sea de prueba bajo `packages/web-component` publica
  todos los paquetes de framework.
- **Canal de compilación compartido.** `scripts/build-packages.mjs` y `tsconfig.release.json` publican
  todos los paquetes de framework. El servidor queda fuera (`global: false`): solo `server/**` lo
  publica.

El nivel de incremento es el nivel más alto de Conventional Commit entre los commits que tocan
archivos publicados en el ámbito del paquete: `!` / `BREAKING CHANGE:` es mayor, `feat` menor, y
cualquier otra cosa parche. Usa los ámbitos `docs(teams): ...`, `fix(server): ...` y similares
(consulta el `AGENTS.md` de la raíz).

## Comprobaciones que se ejecutan para este visor {#checks-that-run-for-this-viewer}

`bun run check:published` (ningún rango `workspace:`, `file:` o `link:` y ningún paquete privado en un
manifiesto publicado; también analiza los archivos `.mjs` y el directorio del servidor) y
`bun run pack:smoke` (instala e importa los tarballs). `bun run test:scripts` ejecuta las pruebas de
scripts propias del visor (`scripts/*.test.mjs`).

El CI ejecuta `bun run pack:smoke --workspace-runtime`: los tarballs se instalan con el `ooxml-core` y
el `ooxml-ui` de esta copia de trabajo (compilados primero), la combinación que se publica junta. Sin
el indicador, `pack:smoke` instala el núcleo y la interfaz del registro, que es lo que obtiene un
usuario externo. El `verify` de la raíz ejecuta estos pasos de paquete solo con `--full`.

Nunca ejecutes `npm publish` ni subas etiquetas de publicación a mano, y nunca edites las versiones
ni los archivos `CHANGELOG.md`; eso lo hace el commit de publicación. Un nombre de paquete nuevo
necesita la configuración inicial manual y el publicador de confianza descritos en la página de la
raíz.
