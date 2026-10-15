# Arquitectura

## Responsabilidades {#ownership}

- `ooxml-core/visio` en el repositorio hermano `ooxml`: ingesta de paquetes, XML, modelo, herencia, geometría en caché y diagnósticos estructurados. Aquí no hay ninguna interfaz de visor.
- `src/controller.ts`: estado del documento y de la vista, selección, eventos, comportamiento de "la última carga gana" y orquestación transaccional de la edición.
- `src/document-history.ts`: historial de ediciones acotado y respaldado por el origen, y estado de exportación de copias VSDX. La mutación del XML y del paquete sigue en el núcleo y se ejecuta mediante un worker dedicado.
- `src/document-text-search.ts`: coincidencia literal acotada sobre el texto visible normalizado de las formas, con resultados inmutables por página. El controlador compartido se encarga de la navegación y de las comprobaciones de reentrada.
- `src/render-svg.ts` y `src/render-text.ts`: presentación SVG y avisos de renderizado.
- `src/export-svg.ts`: serialización estática y acotada de la página actual mediante el mismo renderizador SVG, con recursos rasterizados incrustados y metadatos de compatibilidad.
- `src/print-snapshot.ts`: artefactos de las páginas seleccionadas, sin efectos secundarios, con presupuestos agregados y diagnósticos conservados. No crea marcos de impresión ni implementa políticas de impresora.
- `src/viewer-element.ts`: superficie compartida del navegador.
- `src/contract.ts` y `src/binding.ts`: propiedades, eventos y ciclo de vida mount/update/load/fit/exportSvg/destroy solo de cliente. Las actualizaciones reentrantes más recientes reemplazan el resto del parche anterior.
- Envoltorios de frameworks: solo ciclo de vida del framework y reenvío de eventos y propiedades. Ningún analizador, renderizador ni bifurcación de geometría por framework.

## Límites {#boundaries}

El motor usa actualmente valores en caché de ShapeSheet. Esto no es una evaluación exhaustiva de fórmulas. Las métricas de fuentes del navegador, el ajuste de texto, el espaciado y la sangría mixtos de párrafo, el texto de derecha a izquierda y las viñetas se implementan con trabajo acotado; esto no establece una fidelidad exacta del texto de Visio.

El XML/HTML del documento no se inyecta en el DOM. Los nodos SVG se crean mediante las API del DOM. No se ofrece subida de documentos, telemetría ni obtención de documentos remotos. Las comprobaciones de paquetes y relaciones corresponden al núcleo. El componente del navegador analiza en un worker dedicado, con cancelación y un tiempo límite de 15 segundos en el padre. Presupuestos independientes de escena, metadatos, trabajo de texto y rásteres decodificados protegen el renderizado. Los recursos de imagen se comparten dentro de cada renderizado y se revocan al reemplazarse, desconectarse o destruirse. Los oyentes de carga de fuentes provocan un nuevo ajuste del texto y se eliminan al desconectarse. Estas defensas no son una certificación de seguridad.

El reemplazo experimental de texto sin formato respaldado por el origen admite deshacer y rehacer acotados y la exportación explícita de copias VSDX. El núcleo rechaza las formas vinculadas a plantillas maestras, el texto enriquecido, los campos, los paquetes firmados y el contenido con macros. Los documentos que solo tienen modelo son de solo lectura. Las cachés de fórmulas no se recalculan y la verificación de reapertura en Visio nativo sigue pendiente; esto no es una garantía de viaje de ida y vuelta al editar y guardar. Los métodos del manejador, los eventos y los límites están en la [API del visor](/es/api).

## Seguridad y límites {#safety-and-limits}

El núcleo valida las relaciones OPC, rechaza DTD y entidades y las rutas ZIP sospechosas, y acota la entrada del archivo, el número de entradas, la expansión declarada, la salida real transmitida y la complejidad del XML. El visor comprueba el tamaño del archivo antes de reservar el Blob y valida las escenas suministradas externamente antes de dibujar, incluida la estructura real de los rásteres, el MIME detectado y las dimensiones intrínsecas. La validación de rásteres comprueba la estructura y las sumas de verificación, no la decodificación completa de la entropía; la robustez de los decodificadores nativos sigue siendo relevante. Las relaciones externas nunca se obtienen, las cadenas del documento son nodos de texto y las funciones no admitidas se notifican.

El elemento personalizado del navegador usa un worker de análisis dedicado con cancelación y un límite de 15 segundos en el lado del padre. El controlador sin interfaz y los entornos sin compatibilidad con Worker usan límites cooperativos del núcleo. El renderizado de SVG y de texto sigue ejecutándose en el hilo principal, con límites independientes de escena, ráster, metadatos y trabajo de texto. El proyecto aún no cuenta con certificación de seguridad ni está endurecido para producción.

## Metarchivos incrustados {#embedded-metafiles}

Los metarchivos mejorados incrustados reciben diagnósticos de compatibilidad acotados a nivel de registro. El verificador rechaza los formatos, estilos, estados de asignación y amplificaciones de recursos no admitidos antes de cualquier conversión. No ejecuta datos WMF incrustados, fuentes, imágenes, scripts ni referencias externas. El worker de análisis desechable convierte únicamente un subconjunto estrecho admitido de líneas, rectángulos, elipses y objetos de stock mediante el paquete convertidor publicado, con presupuestos del documento y un plazo máximo controlado por el padre. El análisis directo del núcleo sigue siendo solo de inspección, salvo que se suministre explícitamente un convertidor de confianza en un anfitrión aislado. Los cinco medios EMF reales inspeccionados previamente siguen sin admitirse; la admisión estructural por sí sola no promete renderizado ni fidelidad con Visio.

El núcleo sin interfaz contiene un sanitizador de vectores neutro y un validador de transporte, ambos probados. El worker de análisis usa el convertidor publicado para un subconjunto acotado de primitivas EMF, que comparte el renderizado en vivo, la exportación SVG y la impresión inmutable. El análisis alternativo sin worker sigue sin convertidor. Las pruebas de recorte generadas solo con rutas coinciden con el backend SVG local de Skia; el backend librsvg instalado ignora las intersecciones anidadas de clip-path. Siguen siendo necesarias las comparaciones en el navegador y con Visio nativo. Consulte [la revisión detallada de adopción](research/emf-adoption-review.md).

## Verificación {#verification}

Las pruebas unitarias y de fixtures del núcleo establecen los casos de análisis y modelo que afirman. Las pruebas del visor establecen el ciclo de vida y el comportamiento de la vista. Las pruebas en el navegador establecen la interacción real. Siguen siendo necesarios fixtures creados genuinamente en Visio, renderizados de referencia conocidos y comparaciones medidas para cualquier afirmación de fidelidad. Las pruebas de ejecución de los frameworks son independientes de la compilación de los envoltorios.

## Sitio web {#website}

Compilación estática multipágina con Vite: inicio, guía, registro de capacidades, arquitectura y un área de pruebas independiente. Los enlaces son relativos y los recursos locales. Sin fuentes ni scripts de terceros. Sin paquete del analizador en la página de inicio.

El diseño de la documentación adapta la [referencia de pptx-viewer](https://github.com/ChristopherVR/pptx-viewer). La revisión en profundidad está en `research/pptx-pages-design-review.md`.

## Publicación {#publication}

Siete paquetes de npm exponen la API de Visio sin interfaz y el visor compartido con seis adaptadores nativos de frameworks. Dependen de `ooxml-core/visio` publicado. Las versiones horarias basadas en conventional commits compilan y verifican los artefactos empaquetados antes de la publicación de confianza en npm. GitHub Pages despliega la documentación y la demostración tras las comprobaciones.
