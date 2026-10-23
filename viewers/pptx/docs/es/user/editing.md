---
title: Editar diapositivas
description: Seleccione, mueva, redimensione y dé formato a elementos; inserte texto, formas, imágenes, tablas y gráficos; gestione diapositivas y secciones; deshaga cambios y recupere trabajo sin guardar.
---

# Editar diapositivas

Esta página trata sobre cómo cambiar el contenido de una presentación: trabajar con elementos, dar formato, organizar, gestionar diapositivas y mantener su trabajo a salvo.

::: warning La edición debe estar habilitada
La edición solo está disponible cuando la aplicación está configurada para ello. Si la barra de herramientas muestra una insignia **Solo lectura** y no puede seleccionar ni mover elementos, la aplicación está en modo de solo visualización. Desarrolladores: consulte la documentación de [Props del componente](/react/props).
:::

## Seleccionar elementos

- **Haga clic** en un elemento del lienzo para seleccionarlo. Aparecen controladores a su alrededor.
- **Mayús + clic** en más elementos para añadirlos a la selección.
- **Arrastre sobre el lienzo vacío** para dibujar un recuadro de selección y seleccionar todo lo que contenga.
- Pulse **Ctrl+A** (Cmd en macOS) para seleccionar todos los elementos de la diapositiva.
- Pulse **Escape** para anular la selección.

También puede seleccionar desde la pestaña **Elementos** del inspector, que lista todos los elementos de la diapositiva en orden de apilamiento bajo el encabezado **Orden de capas**. Es útil cuando los elementos se superponen.

![Element selected with resize handles visible](/user-guide/editing-element-selected.jpg)

## Mover, redimensionar y rotar

Con un elemento seleccionado:

- **Mover** - arrástrelo. Aparecen guías de alineación y líneas de ajuste cuando se alinea con la cuadrícula o con otras formas.
- **Desplazar con precisión** - pulse las **teclas de flecha** para pasos pequeños, **Mayús + teclas de flecha** para pasos mayores.
- **Redimensionar** - arrastre cualquier controlador de esquina o de borde.
- **Rotar** - arrastre el controlador de rotación situado sobre el elemento. Mantenga **Mayús** al rotar para ajustar a pasos de 15 grados.
- **Ajustar la forma** - algunas formas muestran un controlador de ajuste adicional que cambia su geometría (por ejemplo, el tamaño de la punta de una flecha).

![Dragging and resizing elements](/user-guide/editing-drag-resize.gif)

Para una colocación exacta, abra la pestaña **Propiedades** del inspector. La tarjeta **Elemento** tiene los campos **X**, **Y**, **An** y **Al**, además de un interruptor de bloqueo que protege el elemento de arrastres accidentales.

::: tip Reglas, cuadrícula y guías
Active **Reglas**, **Cuadrícula**, **Guías** y **Ajustar a la cuadrícula** desde el grupo Mostrar de la pestaña **Vista**. Arrastre una guía para reposicionarla; haga doble clic en una guía para quitarla.
:::

## Editar texto

1. **Haga doble clic** en un cuadro de texto, una forma o una celda de tabla. Aparece un cursor y puede escribir.
2. Pulse **Ctrl+Intro** para confirmar la edición, o **Escape** para cancelarla.

![Inline text editing with cursor active](/user-guide/editing-inline-text.jpg)

Dé formato al texto desde la pestaña **Inicio**:

- **Grupo Fuente** - listas desplegables de familia y tamaño de fuente; **Negrita**, **Cursiva**, **Subrayado**, **Tachado**; **Sombra de texto**; **Aumentar / Disminuir tamaño de fuente**; **Borrar formato**; **Espaciado entre caracteres**; **Cambiar mayúsculas y minúsculas**; **Color de fuente** y **Color de resaltado de texto** (cada uno con muestras predefinidas y un selector de **Color personalizado...**).

  La lista de fuentes se agrupa como lo hace PowerPoint: primero las dos
  **fuentes del tema** de la presentación (marcadas _Encabezados_ y _Cuerpo_), luego las fuentes
  **incrustadas en esta presentación**, luego las que haya
  [añadido en esta sesión](#anadir-una-fuente-que-el-visualizador-no-tiene), y por último la lista
  alfabética completa. Si nada la sobrescribe en el elemento seleccionado, el cuadro
  muestra la fuente que la presentación renderizaría realmente, que dentro de un marcador
  de posición de título es la fuente de encabezados del tema y no una genérica por defecto.

- **Grupo Párrafo** - listas con viñetas y numeradas; controles de sangría; **Alinear a la izquierda / centrar / derecha / Justificar**; **Interlineado**; **Dirección del texto** (horizontal, girado, apilado); **Columnas**.

Los mismos ajustes de texto están disponibles en el panel **Texto** del inspector cuando se selecciona un elemento de texto.

## Insertar elementos

Use la pestaña **Insertar** para añadir contenido a la diapositiva actual:

![Insert tab with element options](/user-guide/editing-insert-tab.jpg)

| Botón                          | Qué inserta                                                                                                                                         |
| ------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Cuadro de texto**            | Un cuadro de texto vacío listo para escribir.                                                                                                       |
| **Forma**                      | Una forma predefinida; elija el tipo de forma en la lista desplegable junto al botón (rectángulos, flechas, llamadas, estrellas, conectores y más). |
| **Imagen**                     | Una imagen de un archivo de su dispositivo.                                                                                                         |
| **Medios**                     | Un archivo de video o audio.                                                                                                                        |
| **Tabla**                      | Una tabla nueva en la diapositiva.                                                                                                                  |
| **Gráfico**                    | Un gráfico; elija **Barras**, **Líneas**, **Circular**, **Anillo**, **Áreas** o **Dispersión** en la lista desplegable junto al botón.              |
| **SmartArt**                   | Un diagrama SmartArt elegido de una galería de diseños.                                                                                             |
| **Ecuación**                   | Una ecuación matemática creada en el editor de ecuaciones.                                                                                          |
| **Acción**                     | Un botón de acción (flechas de navegación, inicio y preajustes similares).                                                                          |
| **Campo**                      | Un campo dinámico: **Número de diapositiva**, **Fecha y hora**, **Encabezado** o **Pie de página**.                                                 |
| **Encabezado y pie de página** | Abre el cuadro de diálogo de encabezado y pie de página de la presentación.                                                                         |

Los elementos nuevos aparecen en la diapositiva seleccionados y listos para moverlos, redimensionarlos y darles estilo.

![Shapes on a slide](/docs-shots/shapes-slide.jpg)

También hay una pestaña **Dibujar** con herramientas a mano alzada: **Seleccionar**, **Lápiz**, **Resaltador**, **Borrador** y **Forma libre**, además de un selector de color de lápiz y un control deslizante de grosor de trazo. La tinta que dibuja se convierte en elementos de la diapositiva.

## Dar formato con el inspector

La pestaña **Propiedades** del inspector cambia según la selección:

| Cuando selecciona...    | Puede cambiar...                                                                                                        |
| ----------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| Nada (la diapositiva)   | Información de la presentación, tema, tamaño de diapositiva, transición de diapositiva y fondo (color o imagen).        |
| Una forma o un conector | Posición, tamaño, rotación, relleno (sólido, degradado, trama, imagen) y trazo.                                         |
| Un elemento de texto    | Fuente, tamaño, color, alineación y espaciado.                                                                          |
| Una imagen              | Recorte y ajustes de imagen.                                                                                            |
| Una tabla               | Estilos de tabla predefinidos, anchos de columna, alturas de fila y formato de celda.                                   |
| Un gráfico              | La cuadrícula de datos de series, tipo de gráfico, ejes, etiquetas de datos, marcadores, líneas de tendencia y colores. |
| SmartArt                | Texto de nodos y diseño.                                                                                                |
| Medios                  | Ajustes de reproducción.                                                                                                |
| Cualquier elemento      | **Animaciones** de entrada, énfasis y salida con ajustes de dirección, temporización y desencadenador.                  |

![Inspector panel showing element properties](/user-guide/editing-inspector-panel.jpg)

### Editar gráficos en el lienzo

Los gráficos se pueden editar directamente sin salir de la diapositiva:

- **Pase el cursor** sobre una marca para ver su valor en una información sobre herramientas.
- **Haga clic** en una barra, punto o sector para seleccionarlo; la celda correspondiente se resalta en la cuadrícula de datos del inspector.
- **Arrastre** una barra o punto hacia arriba o abajo para cambiar su valor (gráficos de barras, líneas, áreas, dispersión y burbujas). Una insignia muestra el valor mientras arrastra; suelte para confirmar o pulse **Escape** para cancelar.
- **Haga doble clic en el título del gráfico** para renombrarlo en el sitio.
- La **hoja de datos** del gráfico, en el inspector, enumera sus categorías y series. Un valor se confirma al pulsar **Intro** o **Tab** o al pasar a otra celda; **Escape** restaura el valor anterior de la celda. Un valor que no es un número se rechaza.

Las marcas de gráficos circulares, de anillo, radiales y apiladas/apiladas al 100 % se seleccionan con un clic; edite sus valores en la cuadrícula de datos. Cada cambio es un único paso de deshacer. Todo esto funciona igual en todos los enlaces (React, Vue, Angular, Svelte y vanilla).

### Editar tablas

Haga clic con el botón derecho en una tabla para ver los comandos de filas y columnas: **Insertar fila arriba / abajo**, **Insertar columna a la izquierda / derecha**, **Eliminar fila**, **Eliminar columna**, **Combinar celdas** y **Dividir celda**. Haga doble clic en una celda para editar su texto.

![A table on a slide](/docs-shots/table-slide.jpg)

## Organizar elementos

![Arrange controls in the ribbon](/user-guide/editing-arrange-tab.jpg)

- **Orden de apilamiento** - haga clic con el botón derecho en un elemento y elija **Traer adelante**, **Enviar atrás**, **Traer al frente** o **Enviar al fondo**. Los mismos comandos están en el grupo Dibujo de la pestaña Inicio, en **Organizar**, y como botones de capa en el inspector.
- **Alinear y distribuir** - con dos o más elementos seleccionados, use los botones de alineación de la pestaña Inicio (izquierda, centro, derecha, arriba, medio, abajo) y **Distribuir** horizontal o verticalmente.
- **Agrupar / Desagrupar** - haga clic con el botón derecho en una multiselección y elija **Agrupar** para que los elementos se muevan y redimensionen juntos; **Desagrupar** los separa de nuevo.
- **Voltear** - use **Voltear H** / **Voltear V** para reflejar la selección.

## Gestionar diapositivas

- **Agregar una diapositiva** - haga clic en **Agregar diapositiva** en la parte inferior del panel de diapositivas, o use **Inicio > Nueva diapositiva**. La flecha junto a Nueva diapositiva abre una galería de los diseños del patrón de la diapositiva actual, cada uno dibujado con su ilustración real y sus marcos de marcador de posición en lugar de listado por nombre. El botón **Diseño** abre la misma galería para cambiar el diseño de la diapositiva actual, con el diseño que usa ahora resaltado con un contorno. Al cambiar de diseño, su contenido pasa a los marcadores de posición correspondientes; lo que el nuevo diseño no tenga dónde alojar permanece en la diapositiva como contenido independiente en lugar de descartarse.
- **Reordenar** - arrastre las miniaturas en el panel de diapositivas, o arrastre las tarjetas en el [clasificador de diapositivas](/es/user/viewing#el-clasificador-de-diapositivas).
- **Duplicar, ocultar o eliminar** - abra el clasificador de diapositivas, haga clic con el botón derecho en una o más diapositivas y elija **Duplicar**, **Ocultar diapositivas** / **Mostrar diapositivas** o **Eliminar**. Copiar y pegar diapositivas completas también funciona allí (**Ctrl+C** / **Ctrl+V**).
- **Secciones** - use **Inicio > Sección** o haga clic con el botón derecho en una miniatura y elija **Añadir sección antes** para agrupar diapositivas bajo encabezados con nombre. Haga clic con el botón derecho en el encabezado de una sección para **Cambiar nombre**, **Eliminar** o **Mover**la.
- **Restablecer** - **Inicio > Restablecer** restaura las posiciones de los marcadores de posición de la diapositiva actual desde su diseño.

## Añadir una fuente que el visualizador no tiene

A veces las presentaciones usan fuentes que no están instaladas en su dispositivo. El
visualizador sustituye la tipografía disponible más parecida, lo que puede cambiar el aspecto
de la presentación. Si tiene el archivo de la fuente, puede entregárselo al visualizador para la
sesión actual:

1. Abra **Archivo > Opciones > General** y busque la sección **Fuentes**.
2. Active **Permitirme añadir archivos de fuente a esta sesión**. El ajuste está desactivado
   por defecto, por lo que el selector de archivos permanece deshabilitado hasta que lo active.
3. Haga clic en **Añadir archivo de fuente** y elija un archivo `.ttf`, `.otf`, `.woff` o `.woff2`.

La familia aparece en **Fuentes añadidas en esta sesión** y en la lista de fuentes de la pestaña
Inicio, y las diapositivas que la usan se vuelven a renderizar con la tipografía real. Los archivos cuyos
nombres indican un grosor o estilo (`Inter-SemiBoldItalic.woff2`) se registran solo bajo el nombre de la
familia, de modo que los cuatro archivos de una misma familia se combinan en una única entrada sobre la
que actúan los botones **Negrita** y _Cursiva_.

Las fuentes añadidas se mantienen solo en memoria. Nunca se suben, nunca se escriben en la
presentación y se descartan al recargar la página, por lo que un archivo de fuente con licencia
permanece donde usted lo puso. Para enviar una fuente _con_ la presentación, use
**Archivo > Incrustar fuentes**.

## Editar las notas del orador

Abra el panel de notas con el botón **Notas** de la barra de estado y escriba. La barra de herramientas de notas ofrece **Negrita**, **Cursiva**, **Subrayado**, **Tachado**, listas con viñetas y numeradas, controles de sangría e inserción de enlaces, además de un interruptor entre el editor de texto enriquecido y el de texto sin formato. Las notas se guardan con la presentación y aparecen en la [vista del presentador](/es/user/presenting#vista-del-presentador).

## Buscar y reemplazar

1. Pulse **Ctrl+F**, o use **Inicio > Edición > Buscar** / **Reemplazar**.
2. Escriba el texto que desea buscar. El contador **x de y** muestra las coincidencias en todas las diapositivas; recórralas con los botones anterior / siguiente.
3. Active **Coincidir mayúsculas y minúsculas** si es necesario.
4. Para cambiar texto, escriba un reemplazo y haga clic en **Reemplazar** o **Reemplazar todo**.

## Comentarios

- Para comentar un elemento, haga clic con el botón derecho en él y elija **Agregar comentario**; o abra la pestaña **Comentarios** del inspector y use el cuadro **Agregar comentario** de la diapositiva actual.
- Cada hilo admite **Responder**, **Editar**, **Resolver** (y **Reabrir**) y **Eliminar**.
- Las diapositivas con comentarios muestran una insignia de comentario en su miniatura. El botón **Comentarios** de la pestaña **Revisar** activa o desactiva el panel.

## Deshacer y rehacer

- **Deshacer** - **Ctrl+Z**, o el botón de flecha de la barra de título.
- **Rehacer** - **Ctrl+Mayús+Z** o **Ctrl+Y**.

Cada edición se captura como un paso por el que puede avanzar y retroceder. Las acciones continuas, como arrastrar, cuentan como un solo paso.

## Cortar, copiar, pegar, duplicar

| Acción   | Atajo                    |
| -------- | ------------------------ |
| Copiar   | **Ctrl+C**               |
| Cortar   | **Ctrl+X**               |
| Pegar    | **Ctrl+V**               |
| Duplicar | **Ctrl+D**               |
| Eliminar | **Supr** o **Retroceso** |

Los mismos comandos aparecen al hacer clic con el botón derecho en un elemento. Consulte la lista completa en la página de [Atajos de teclado](/es/user/shortcuts).

Para pegar una imagen desde otra aplicación, copie la propia imagen, haga clic
en el lienzo de la diapositiva y pulse **Ctrl+V** (**Cmd+V** en macOS). Cuando el navegador
entrega un archivo de imagen, se inserta en el centro de la diapositiva, ajustado a sus
límites y seleccionado. La inserción es un único paso de deshacer. Si pega mientras edita
texto o una celda de tabla, el contenido permanece en ese editor y no inserta una imagen
independiente en la diapositiva. Este atajo no importa solo URL de imágenes ni HTML.

El portapapeles interno de elementos del editor conserva la prioridad. Si ha copiado
elementos de diapositiva en este visualizador, **Ctrl+V** sigue pegando esos elementos.
En ese caso, use **Insertar > Imagen** para elegir un archivo de imagen externo.

## Autoguardado y recuperación

- El interruptor de **Autoguardado** de la barra de título activa o desactiva las instantáneas periódicas. Cuando está activado, la aplicación guarda una copia de recuperación de su trabajo en el almacenamiento IndexedDB del navegador. El primer cambio sin guardar inicia un temporizador, y la copia se escribe cuando termina (cada un par de minutos por defecto). Mientras no cambie nada, no se escribe nada.
- La barra de estado muestra el estado actual: **Cambios sin guardar**, **Guardando...**, **Guardado [hora]** o **Todo guardado**.
- Si cierra la pestaña con cambios sin guardar y vuelve, un aviso titulado **Cambios sin guardar recuperados** ofrece **Restaurar** y **Descartar**.
- El **Historial de versiones** (en el menú de desbordamiento **"..."**) lista la instantánea de recuperación almacenada con su hora y tamaño, y permite restaurarla o eliminarla.

::: warning El autoguardado no es guardar un archivo
El autoguardado conserva una copia de recuperación solo dentro de su navegador. Para producir un archivo `.pptx` real, use el botón **Guardar** de la barra de título o **Archivo > Guardar**; el archivo se descarga en su dispositivo. Consulte [Exportación](/es/user/exporting).
:::

## A continuación

- Presente su presentación: [Presentar](/es/user/presenting)
- Guárdela y compártala: [Exportación](/es/user/exporting)
- Trabaje más rápido: [Atajos de teclado](/es/user/shortcuts)
