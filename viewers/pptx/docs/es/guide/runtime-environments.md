---
title: Entornos de ejecución
description: Dónde se ejecuta pptx-viewer (navegador, Node.js, Web Worker) y el comportamiento específico de cada plataforma que se deriva de ejecutarse dentro del entorno aislado (sandbox) de un navegador.
---

# Entornos de ejecución

`pptx-viewer` se divide en un motor principal sin DOM y en enlaces de interfaz que solo funcionan en el navegador. Esta página explica dónde se ejecuta cada parte, y el comportamiento que se deriva del entorno aislado del navegador y no de una función que falte.

## Dónde se ejecuta {#where-it-runs}

| Entorno                  | Funciona    | Limitaciones                                                                                                                                                                                                                     |
| ------------------------ | ----------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Navegador                | Sí          | El conjunto completo de funciones: análisis, renderizado, edición, exportación y colaboración.                                                                                                                                   |
| Node.js (y sin servidor) | Solo núcleo | `pptx-viewer-core` (cargar, editar, guardar, conversión a Markdown/SVG, cifrado) no depende del DOM. Los enlaces de interfaz, la exportación de rasterizado (`html2canvas`) y la conversión EMF/WMF son funciones del navegador. |
| Web Worker               | Solo núcleo | Mismo alcance que Node.js: el motor no depende del DOM.                                                                                                                                                                          |

## Notas por plataforma {#platform-notes}

Estas no son carencias que haya que subsanar: se derivan directamente de ejecutarse dentro del entorno de seguridad aislado de un navegador, que no puede hacer lo que sí puede la aplicación de escritorio.

### Fuentes {#fonts}

Un navegador no tiene un catálogo de fuentes propio sobre el que dibujar, como sí ocurre en la aplicación de escritorio, así que `pptx-viewer` resuelve cada fuente de la presentación en tres pasos: usarla tal cual si está instalada en el equipo del lector; si no, descargar antes del diseño su clon verificado compatible en métricas desde Google Fonts (Calibri -> Carlito, Cambria -> Caladea, Arial -> Arimo, Times New Roman -> Tinos, Courier New -> Cousine, Georgia -> Gelasio, de modo que los saltos de línea y las dimensiones del texto sigan coincidiendo con PowerPoint); y solo si no existe ninguna de las dos, recurrir, mediante la clasificación PANOSE, a una fuente genérica de la misma clase, lo que puede desplazar las métricas.

Mientras se carga una presentación, el visualizador espera brevemente (150 ms) a las fuentes de Google Fonts que dibujan sus diapositivas, para que las fuentes que ya están en caché nunca aparezcan primero como sustitutos. Una primera descarga en una red lenta no retiene la presentación: pasados 150 ms, o de inmediato si la petición falla (sin conexión, bloqueada), la presentación se muestra con sustitutos compatibles en métricas, la barra de estado indica "Downloading fonts (1 of 2)" mientras dura la descarga, y una fuente que llega tarde se vuelve a diseñar cuando llega. Las fuentes instaladas e incrustadas no necesitan espera.

### Fuentes y privacidad {#fonts-and-privacy}

Obtener un clon desde Google Fonts significa que el navegador del lector contacta con `fonts.googleapis.com` y `fonts.gstatic.com`, lo que revela a Google la dirección IP del lector. Por defecto, el lector decide con la opción "Download fonts that are not installed or embedded (Google Fonts)" de File > Options > Trust Center (Archivo > Opciones > Centro de confianza), que está activada. Desactivarla surte efecto de inmediato en la presentación abierta (el visualizador quita su hoja de estilos de webfonts) y en todas las cargas posteriores, y la elección se guarda junto con el resto de opciones. Un host puede decidir en su lugar con la opción `googleFonts`, disponible en todos los enlaces:

- sin definir (el valor por defecto): decide la opción del Centro de confianza; si está activada, los clones se cargan desde Google Fonts como se ha descrito antes;
- `googleFonts: false`: ninguna petición va a Google; el texto usa las fuentes instaladas, las fuentes incrustadas en la presentación o la sustitución PANOSE, así que los saltos de línea pueden diferir de PowerPoint. La opción del Centro de confianza aparece desactivada y de solo lectura;
- `googleFonts: { baseUrl: 'https://fonts.example.com/css2' }`: la misma petición CSS2 va a un espejo autoalojado, y la opción del Centro de confianza aparece activada y de solo lectura. El espejo debe responder al formato de consulta CSS2 de Google Fonts (`?family=...&display=swap`) con CSS `@font-face` compatible, y sus archivos de fuente también deben estar autoalojados para que ninguna petición llegue a Google. Un `baseUrl` que no sea una URL `http(s)` absoluta desactiva la alternativa (la opción aparece desactivada y de solo lectura) y registra una advertencia.

La opción también rige la descarga de los contornos de glifos de WordArt y las webfonts incrustadas en las exportaciones de imagen de las diapositivas, porque ambas reutilizan la URL de hoja de estilos que resolvió el visualizador. Las dos leen las hojas CSS y los archivos de fuente con `fetch()`, así que un espejo debe enviar cabeceras CORS (`Access-Control-Allow-Origin`) para que funcionen, como hacen los propios servidores de Google.

Cómo elige el visualizador una variante de fuente para cada tramo, y cómo espera las descargas, se describe en la [carga de fuentes](https://github.com/ChristopherVR/ooxml/blob/main/docs/font-loading.md). Todas las peticiones que pueden hacer los paquetes aparecen en [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

### Contenido externo y privacidad {#external-content-and-privacy}

Una presentación puede enlazar imágenes (incluidos los rellenos de imagen y las viñetas con imagen), vídeos, audio e imágenes de póster de fotogramas a otros servidores (relaciones `TargetMode="External"`) en lugar de incrustarlos. Al abrir una presentación así, el navegador del lector descarga esos recursos, lo que indica al responsable de ese servidor la dirección IP del lector y que la presentación se ha abierto. Por defecto, decide la opción "Allow external content" del Centro de confianza, que está activada. Un host puede decidir en su lugar con la opción `allowExternalContent`, disponible en todos los enlaces:

- sin definir (el valor por defecto): decide la opción del Centro de confianza;
- `allowExternalContent: false`: no se descarga ningún contenido enlazado, así que las diapositivas muestran solo contenido incrustado, y la opción del Centro de confianza aparece desactivada y de solo lectura;
- `allowExternalContent: true`: el contenido enlazado se descarga, y la opción aparece activada y de solo lectura.

El valor se aplica a todas las presentaciones que abre el visualizador, incluidas las abiertas para Comparar y las diapositivas reutilizadas de otro archivo; un cambio surte efecto desde la siguiente carga. Los hipervínculos no se ven afectados: solo se abren cuando el lector los sigue. Las fuentes son otro asunto: las fuentes de la presentación que no están instaladas ni incrustadas se cargan desde Google Fonts, salvo que la opción del Centro de confianza "Download fonts that are not installed or embedded" o la opción `googleFonts` lo desactiven (consulte [Fuentes y privacidad](#fonts-and-privacy)).

En una sesión de colaboración, las diapositivas llegan de los pares ya analizadas, y el ajuste no se les vuelve a aplicar: un par cuyo visualizador permite el contenido externo puede compartir una diapositiva que haga referencia a una URL enlazada. Si esto importa, use el mismo valor de `allowExternalContent` en todos los participantes.

### Límites de carga {#load-limits}

Un host que abre presentaciones que no son de confianza (subidas de usuarios, adjuntos) puede acotar el trabajo que puede causar una presentación con la opción `limits`, disponible en todos los enlaces y pasada directamente a `PptxHandler.load`:

| Campo                  | Valor por defecto | Limita                                                                                                                           |
| ---------------------- | ----------------- | -------------------------------------------------------------------------------------------------------------------------------- |
| `maxInputBytes`        | sin límite        | El archivo tal como se entrega al visualizador, comprobado antes de abrirlo.                                                     |
| `maxUncompressedBytes` | 500 MiB           | El archivo una vez descomprimido, comprobado a partir de los tamaños declarados y de nuevo mientras se lee cada parte.           |
| `maxEntries`           | 65.536            | Las partes (entradas ZIP) del archivo; 65.536 es también el valor más alto aceptado.                                             |
| `maxSlides`            | sin límite        | Las diapositivas listadas en `ppt/presentation.xml`, comprobadas antes de analizar ninguna diapositiva.                          |
| `maxShapes`            | sin límite        | Los elementos creados en las diapositivas (incluidos los miembros de grupos), contados a medida que se analiza cada diapositiva. |

Si no se define el campo, o si falta, se mantiene el valor por defecto, así que un visualizador sin `limits` se comporta como antes. Una presentación que supere un límite deja de cargarse en esa comprobación, y el visualizador muestra su pantalla de error de carga con un mensaje localizado que nombra el límite y el valor de la presentación, por ejemplo "Esta presentación tiene 120 diapositivas, más de las 100 que acepta este visualizador." Un valor que no sea un número finito positivo se rechaza como error en lugar de ignorarse. Los límites se aplican a la presentación que abre el visualizador (la `content` / `source` y File > Open, es decir, Abrir), no a un archivo que el usuario elija para Comparar o Reutilizar diapositivas; un cambio surte efecto desde la siguiente carga. Sin el visualizador, pase el mismo objeto como `handler.load(buffer, { limits })` (consulte [Opciones de carga](/core/loading#load-options)).

### Ejecutar acciones de programa (`ppaction://program`) {#run-program-actions-ppaction-program}

El ajuste de acción "Run program" de PowerPoint se analiza y se vuelve a escribir sin pérdidas (la ruta o el comando del programa que se escribió sobrevive a la carga y al guardado sin cambios), y el cuadro de diálogo Action Settings permite verlo y editarlo como cualquier otra acción mientras se edita. Lo que un navegador no puede hacer es lo que hace PowerPoint cuando la acción se activa durante una presentación: lanzar un ejecutable local. No existe ninguna API de la plataforma web para ello, por diseño (una página que pudiera lanzar programas locales arbitrarios sería un agujero de seguridad crítico).

En su lugar, al hacer clic en una forma de Run Program durante una presentación en marcha, se muestra un aviso no bloqueante con el comando exacto resuelto (la ruta y los argumentos que escribió el autor; OOXML los guarda como una única cadena opaca, y no hay nada que dividir) junto con un botón Copy de un solo clic. Así, el presentador ve lo que PowerPoint habría ejecutado y puede lanzarlo él mismo fuera del navegador si lo desea. El clic sigue contando como gestionado (la presentación tampoco avanza), igual que con cualquier otro tipo de acción. Está implementado una sola vez como función de decisión compartida (la intención `runProgram` de `src/ui/src/pptx/render/presentation-action.ts`, junto con `src/ui/src/pptx/render/run-program-notice.ts`), que usan por igual los cinco enlaces, y `e2e/pptx/run-program-notice.spec.ts` prueba el aviso, el botón Copy y el consumo del clic en los cinco demos.

Durante su desarrollo se detectaron y corrigieron dos errores reales. React mostraba el aviso como hermano del elemento sobre el que la presentación en marcha llama a `requestFullscreen()`, y la capa superior de un elemento en pantalla completa a nivel del sistema operativo se pinta por encima de todo lo que queda fuera de él, sin importar el `z-index` de CSS; el botón Copy era visible pero no se podía pulsar. Ahora se renderiza a través de la ranura de superposición del propio escenario de presentación, junto con el resto de elementos de interfaz que solo aparecen durante la presentación. Por otro lado, una ruta de destino que no tiene forma de URI (una ruta escrita como `notepad.exe C:\temp\notes.txt`) no se reconocía como destino de relación externa al guardar, así que el paquete escrito fallaba la validación de `TargetMode="External"` para `ppaction://program`, `hlinkfile` y `hlinkpres` por igual; la ruta de guardado ahora fuerza `External` en los tres casos.

### Reproducción de multimedia {#media-playback}

La reproducción de audio y vídeo depende del soporte de códecs del propio navegador: WMV y otros códecs antiguos pueden no reproducirse, y el contenido multimedia protegido por DRM no se reproducirá. Es un hecho de la plataforma, no una función que falte: `pptx-viewer` delega la reproducción en el elemento multimedia nativo del navegador, y un navegador solo decodifica lo que incluye un códec para ello.

## Lectura relacionada {#related-reading}

- [Limitaciones](/es/guide/limitations) - brechas abiertas y sin resolver.
- [Fidelidad de efectos visuales](/es/guide/visual-effects) - aproximaciones de efectos en CSS/SVG y la evidencia que las respalda.
