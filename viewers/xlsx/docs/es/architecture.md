# Arquitectura y responsabilidades

## Dirección de las dependencias

```text
published, one per framework (self-contained bundles)
  @christophervr/xlsx-react-viewer | xlsx-vue-viewer | xlsx-angular-viewer
  xlsx-svelte-viewer | xlsx-solid-viewer | xlsx-vanilla-viewer
  ----------------------------------------------------------------------------
  internal, private, inlined into every bundle:
        bindings (lifecycle and events only)
                         |
               <xlsx-editor> web component (grid painter, ribbon, dialogs)
                         |
        ooxml-core (external, published)                 ooxml-ui (external)
   /xlsx       model, formula engine, layout, edit,      office-ui-* controls
               read/write .xlsx
   /xlsx/load  format detection, .xls (ole2 codecs inlined), CSV
        ^
   @christophervr/xlsx-core (external): thin re-export of ooxml-core/xlsx and /xlsx/load
```

Solo se publican `@christophervr/xlsx-core` y los seis paquetes de framework. `web-component` y `bindings` son paquetes de workspace `private`: son el código fuente compartido del editor, nunca un destino de instalación en npm. `scripts/build-packages.mjs` los empaqueta dentro de cada paquete de framework, de modo que un tarball solo importa `@christophervr/xlsx-core`, `ooxml-core`, `ooxml-ui` y su dependencia par del framework. `scripts/check-published-refs.mjs` y la prueba de humo del empaquetado hacen fallar la compilación si un paquete interno o `ole2` se filtra en un tarball.

## El núcleo es responsable de la lógica

Todo lo que no es interfaz de usuario vive en el área `xlsx` de `ooxml-core` (código fuente en el repositorio `ChristopherVR/ooxml`):

- el modelo de libro (`Workbook`, `Worksheet`, `Cell`, estilos de celda resueltos y sin duplicados),
- direcciones y rangos (`parseAddress`, `formatRange`, ...),
- formatos numéricos (`formatValue`, `parseCellInput`),
- el analizador de fórmulas y el motor de cálculo (`createCalcEngine`, seguimiento de dependencias, matrices dinámicas),
- comandos de edición con deshacer y rehacer (`createEditSession`),
- lógica de vista sin DOM: métricas de la cuadrícula, celdas visibles, vistas de celda, formatos condicionales, modelos de vista de gráficos,
- el lector y escritor de `.xlsx` y, en `/xlsx/load`, la detección de formato, el lector de `.xls` heredado (mediante los códecs compartidos de ole2) y CSV.

Un libro cargado conserva en `workbook.source` las partes que el modelo no representa, de modo que al guardar se transmiten intactas.

## El componente web dibuja

`packages/web-component` es la única interfaz de edición: crea una sesión de edición sobre el libro, dibuja lo que devuelven las funciones de diseño del núcleo, dirige la entrada del teclado, el ratón y la cinta de opciones hacia los comandos de edición, y se encarga de los estilos, las configuraciones regionales y los cuadros de diálogo. El Shadow DOM acota sus estilos; el registro solo se produce en el navegador y es idempotente. Las pruebas de navegador ejercitan el mismo contrato a través de cada enlace.

`packages/bindings` se encarga únicamente de la integración con el framework: montaje y desmontaje, actualización de propiedades, reenvío de eventos y manejadores imperativos. Las nuevas funciones del editor pertenecen al componente web (o al núcleo); un cambio en un enlace solo se justifica por un requisito del ciclo de vida o de la API del framework.

## Límite de compilación y publicación

`@christophervr/xlsx-core` sigue siendo una dependencia real de cada paquete de framework en lugar de empaquetarse dentro: las aplicaciones importan de él directamente `createWorkbook` y los tipos del modelo, de modo que una única copia compartida mantiene `Workbook` idéntico en ambos lados. `ooxml-core` y `ooxml-ui` también son dependencias reales. Todo lo demás que es interno se empaqueta; ole2 está integrado dentro del núcleo y no es una dependencia de nada aquí. Consulte [publicación](/releasing).
