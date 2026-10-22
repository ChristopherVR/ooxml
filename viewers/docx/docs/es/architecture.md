# Arquitectura y responsabilidades

## Dirección de las dependencias

```text
published, one per framework (self-contained bundles)
  docx-react-viewer | -vue-viewer | -angular-viewer | -svelte-viewer | -solid-viewer | -vanilla-viewer
  ------------------------------------------------------------------------------------------------
  internal, private, inlined into every bundle:
        bindings (lifecycle and events only)
                         |
                  web-component editor (canvas text measurer)
                         |
        ooxml-core (external, published)
   /docx/layout   /docx/load (format detection, legacy .doc;
   /collab        ole2 codecs inlined by the core)
        ^
   docx-core (external): thin re-export of ooxml-core/docx and /docx/embedded
```

Solo se publican `docx-core` y los seis paquetes `*-viewer`. `web-component` y `bindings` son paquetes `private` del espacio de trabajo: son el código fuente compartido del editor, nunca un destino de instalación en npm. `scripts/build-packages.mjs` los empaqueta dentro de cada paquete de framework, de modo que un tarball solo importa `docx-core`, `ooxml-core`, las bibliotecas de ProseMirror y sus dependencias de framework (peers). `scripts/check-published-refs.mjs` y la prueba de humo del empaquetado hacen fallar la compilación si un paquete interno o `ole2` se filtra a un tarball.

El modelo de documento canónico, el analizador de DOCX y el serializador residen en el área `docx` del paquete `ooxml-core`; `docx-core` es un punto de entrada publicado y muy delgado que lo reexporta. El modelo editable es una proyección deliberadamente pequeña de OOXML. Una sesión cargada es propietaria del paquete original y del modelo de referencia, de modo que guardar ediciones compatibles puede conservar las partes ajenas a esa proyección. Un guardado sin cambios devuelve los bytes originales.

`ooxml-core/docx/load` detecta las firmas del contenedor y selecciona DOCX o DOC heredado. Su lector heredado extrae los párrafos del cuerpo principal de archivos de Word 97-2003, usa los códecs de ole2 (integrados en el núcleo) para CFB y las estructuras binarias de Word, y traduce el texto extraído al modelo común. Su vía de edición solo acepta texto de párrafo simple cuando se superan las comprobaciones de seguridad. `ooxml-core/docx/layout` es el motor de paginación sin DOM; el editor le inyecta su `TextMeasurer` basado en canvas. Los asistentes de colaboración en tiempo real (validación, identidad, reglas de ordenación) proceden de `ooxml-core/collab`; el cliente y la autoridad de `prosemirror-collab`, el complemento de presencia y las decoraciones permanecen en el editor (`ooxml-ui/docx`).

El elemento `<docx-editor>` (código en `ooxml-ui/docx`, que `web-component` reexporta) es la única interfaz de edición. Convierte el modelo en ProseMirror, aplica las transacciones de edición y repara las identidades de bloque, emite los cambios del modelo y llama a la función de guardado de la sesión cargada. La superficie de comandos actual admite la edición de párrafos, el formato directo básico del texto, la alineación, tablas simples y el zoom sobre una superficie de edición continua. El diseño de impresión es una vista paginada aproximada de `ooxml-core/docx/layout`, no la paginación de Word. El Shadow DOM acota los estilos de la barra de herramientas y de la página. El registro solo se produce en el navegador y es idempotente. Las pruebas de navegador ejercitan el mismo contrato a través de cada adaptador.

`bindings` se ocupa únicamente de la integración con los frameworks: montaje y desmontaje, actualización de propiedades, reenvío de eventos y referencias imperativas. Las nuevas funciones del editor pertenecen al componente web; solo se justifica modificar un adaptador por un requisito del ciclo de vida o de la API del framework.

## Código compartido entre los dos visores

El lector y escritor de CFB y los detalles internos binarios de Word son canónicos en el paquete `ole2` (`@christophervr/ole2`). El modelo DOCX, el analizador, el escritor y la API de DOCX incrustado (`docx-core/embedded`) son canónicos en `ooxml-core/docx` (reexportado por `docx-core`). Ninguna implementación moderna de DOCX pertenece a `ole2`.

El paquete compartido `ole2` no depende de React ni de ninguna interfaz de Word o PowerPoint, y se ocupa únicamente de CFB y de los formatos binarios heredados de Word. DOCX es OOXML y pertenece a `ooxml-core/docx`.

## Límite de compilación y publicación

Se publican siete paquetes (`docx-core` y los seis paquetes de framework; consulte [publicación](/releasing)), con versiones automatizadas en npm. El paquete CFB compartido no contiene ningún códec de DOCX.

`docx-core` sigue siendo una dependencia real de cada paquete de framework en lugar de ir empaquetado: las aplicaciones importan `createDocument` y los tipos del modelo directamente desde él, de modo que una única copia compartida mantiene idéntico el `DocumentModel` a ambos lados. Las bibliotecas de ProseMirror también son dependencias reales, no empaquetadas, para que un gestor de paquetes pueda deduplicar una sola copia para el editor y para cualquier otro código ProseMirror de la aplicación. `ooxml-core` es igualmente una dependencia real (aporta el diseño, la carga y la colaboración). Todo lo demás que es interno se empaqueta para que los consumidores no necesiten copias de código fuente hermanas ni paquetes sin publicar; ole2 está integrado dentro del núcleo y no es una dependencia de nada de lo que hay aquí.

Revise la hoja de ruta de compatibilidad antes de tratar esto como equivalencia de funciones con Word. Que las compilaciones de paquetes y las rutas de demo funcionen no demuestra equivalencia en renderizado, paginación ni ida y vuelta con Word.
