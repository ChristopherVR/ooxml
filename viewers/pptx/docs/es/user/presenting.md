---
title: Presentar
description: Inicie una presentación a pantalla completa, contrólela con el teclado, use la vista del presentador, anote con lápiz y resaltador y ensaye los tiempos.
---

# Presentar

El modo de presentación convierte el visualizador en una presentación a pantalla completa que reproduce sus transiciones y animaciones, con vista del presentador, temporizadores y herramientas de anotación en vivo.

## Iniciar una presentación

Desde la pestaña **Presentación con diapositivas**:

1. Haga clic en **Desde el principio** para empezar en la diapositiva 1, o en **Desde la diapositiva actual** para empezar donde está.
2. La diapositiva se expande para llenar la pantalla y los elementos de edición desaparecen.

También puede hacer clic en el botón **Presentar** de la fila superior de la barra de herramientas, o en el botón **Presentación con diapositivas** de la barra de estado. La flecha junto a **Presentar** abre un menú con **Vista del presentador**, **Ensayar intervalos**, **Configurar presentación**, **Presente en línea** y un interruptor de **Subtítulos**.

![Slideshow mode fullscreen](/user-guide/presenting-slideshow.jpg)

**Configurar presentación** controla opciones de reproducción como **Usar tiempos, si están presentes** (avance automático con los intervalos de diapositiva guardados) y repetir la presentación en bucle hasta que pulse Escape.

## Controlar la presentación

| Para...                                                       | Haga esto                                                                                   |
| ------------------------------------------------------------- | ------------------------------------------------------------------------------------------- |
| Diapositiva siguiente / siguiente paso de animación           | Haga clic, o pulse **Flecha derecha**, **Av Pág**, **Barra espaciadora**, **Intro** o **N** |
| Diapositiva anterior                                          | Pulse **Flecha izquierda**, **Re Pág**, **Retroceso** o **P**                               |
| Saltar a una diapositiva concreta                             | Escriba su número y pulse **Intro**                                                         |
| Primera / última diapositiva                                  | Pulse **Inicio** / **Fin**                                                                  |
| Poner la pantalla en negro                                    | Pulse **B** (vuelva a pulsar para reanudar)                                                 |
| Poner la pantalla en blanco                                   | Pulse **W** (vuelva a pulsar para reanudar)                                                 |
| Alternar la vista del presentador                             | Use el botón **Vista del presentador** de la barra de herramientas (sin atajo de teclado)   |
| Mostrar u ocultar la barra de herramientas de la presentación | Pulse **Ctrl+H**                                                                            |
| Mostrar todas las diapositivas                                | Pulse **Ctrl+S**                                                                            |
| Terminar la presentación                                      | Pulse **Escape**                                                                            |

### La barra de herramientas de la presentación

Una barra de herramientas flotante se sitúa en la parte inferior de la pantalla con, de izquierda a derecha: **Diapositiva anterior**, un contador de diapositivas (por ejemplo `3 / 12`), **Diapositiva siguiente**, un temporizador de tiempo transcurrido, **Puntero láser**, **Lápiz**, **Resaltador**, **Borrador**, **Borrar anotaciones**, **Vista del presentador** y **Finalizar presentación**.

![Slideshow toolbar](/user-guide/presenting-toolbar.jpg)

::: tip Ocultación automática
La barra de herramientas se desvanece tras unos 3 segundos sin movimiento del ratón. Mueva el ratón (sobre todo hacia la parte inferior de la pantalla) para que reaparezca, o pulse **Ctrl+M** para fijarla o soltarla.
:::

## Transiciones y animaciones

Cuando avanza entre diapositivas, se reproduce la **transición** asignada a la diapositiva entrante: desvanecimientos, empujes, barridos, revelados, morph y más. Retroceder es instantáneo.

![Slideshow with transitions between slides](/user-guide/presenting-transitions.gif)

Las **animaciones** se reproducen según sus desencadenadores:

- **Al hacer clic** - cada clic o tecla de avance reproduce el siguiente paso de animación; la diapositiva solo cambia cuando se han reproducido todos los pasos.
- **Con la anterior** / **Después de la anterior** - estos pasos se encadenan automáticamente tras el que los precede.

Las diapositivas con intervalos guardados avanzan automáticamente cuando **Usar tiempos, si están presentes** está activado en Configurar presentación. Los elementos con desencadenadores de desplazamiento o clic (incluidos los botones de acción y los hipervínculos) responden durante la presentación.

::: warning Notas de fidelidad
Las transiciones y animaciones son aproximaciones del comportamiento de PowerPoint. Morph interpola la posición, el tamaño, la opacidad, la rotación y el color de los elementos coincidentes, y también transforma entre geometrías de forma distintas: cuando un elemento coincidente cambia de tipo de forma (o de contorno dirigido por ajustes) entre las dos diapositivas, su contorno se resuelve en un polígono y se interpola vértice a vértice hacia el contorno de la forma de destino, en lugar de fundirse. Consulte [Limitaciones](/es/guide/limitations) para otras aproximaciones de transiciones y animaciones.

**Los elementos sin pareja se funden, igual que en PowerPoint.** Los elementos se emparejan entre las dos diapositivas mediante nombres explícitos con `!!`, la identidad `a16:creationId` / id de forma nativo y la proximidad entre elementos del mismo tipo; un nombre simple del Panel de selección sin prefijo `!!` tampoco es una señal de emparejamiento que respete el propio Morph de PowerPoint, así que dar el mismo nombre a dos formas no hace que se transformen una en otra. Un elemento sin pareja en la diapositiva siguiente se desvanece en lugar de transformarse. Confirmado con PowerPoint 2016 real (COM `CreateVideo`, 30 fps, una presentación de dos diapositivas con un elemento sin pareja solo en la diapositiva 1): el elemento se mantiene a plena opacidad hasta el último fotograma del tiempo de permanencia de la diapositiva 1, y luego se desvanece suavemente hasta el fondo en 10 fotogramas (aproximadamente 0,45 s de los 2 s de la transición creada), quedando totalmente transparente, en lugar de desaparecer de golpe en un solo fotograma.
:::

## Vista del presentador

La vista del presentador le muestra información que el público no ve:

- La diapositiva actual (con sus herramientas de anotación todavía disponibles).
- Una vista previa de la **Diapositiva siguiente**.
- **Notas del orador**, con controles de tamaño de fuente; puede editar las notas allí mismo.
- La hora actual, un temporizador transcurrido con pausa y restablecimiento, y una barra de progreso.

Active la vista del presentador con el botón **Vista del presentador** de la barra de herramientas de la presentación, o desde **Presentación con diapositivas > Vista del presentador** en la cinta. No tiene atajo de teclado: `N` es la tecla de "diapositiva siguiente" de PowerPoint.

![Presenter view showing notes, next slide preview, and timer](/user-guide/presenting-presenter-view.jpg)

La barra de herramientas superior de la consola del presentador también ofrece: pausar/restablecer el temporizador, una cuadrícula de **Diapositivas** para saltar a cualquiera, controles de zoom, las herramientas de puntero, botones de oscurecer **B** / **W**, un interruptor de subtítulos y un botón de **Pantalla del público** que abre la presentación en una segunda ventana; use **Intercambiar** para cambiar qué pantalla muestra qué, de modo que su público vea solo las diapositivas mientras usted conserva la consola.

## Anotar en vivo

Dibuje sobre las diapositivas mientras presenta para enfatizar puntos:

| Herramienta                 | Qué hace                                                | Atajo                        |
| --------------------------- | ------------------------------------------------------- | ---------------------------- |
| **Flecha**                  | El puntero normal; no deja marcas.                      | **Ctrl+A**                   |
| **Puntero láser**           | Un punto brillante que sigue al cursor; no deja marcas. | **Ctrl+L**                   |
| **Lápiz**                   | Tinta a mano alzada del color que elija.                | **Ctrl+P**                   |
| **Resaltador**              | Trazos de resaltado translúcidos.                       | (solo barra de herramientas) |
| **Borrador**                | Elimina trazos de tinta individuales.                   | **Ctrl+E**                   |
| **Borrar anotaciones**      | Elimina toda la tinta de la diapositiva actual.         | **E**                        |
| **Mostrar / ocultar tinta** | Oculta la tinta sin eliminarla.                         | **Ctrl+M**                   |

::: tip Por qué las herramientas son combinaciones con Ctrl
PowerPoint reserva las letras sueltas para la navegación y los comandos de pantalla (`N` siguiente,
`P` anterior, `B` negro, `W` blanco, `E` borrar todo), por lo que las herramientas de anotación usan la
combinación con **Ctrl**. Este visualizador sigue la misma división.
:::

Para cambiar el color, haga clic con el botón derecho en el botón del lápiz o del resaltador (o en la pequeña flecha junto a él) y elija en la cuadrícula de muestras. Elegir un color también activa esa herramienta.

![Pen annotation on a slide during presentation](/user-guide/presenting-pen-annotation.gif)

Las anotaciones se conservan por diapositiva: pase a otra diapositiva y vuelva, y su tinta seguirá allí.

![Annotation drawn on a slide](/user-guide/presenting-annotation.jpg)

Cuando termina una presentación que tiene tinta, un cuadro de diálogo **¿Mantener anotaciones?** pregunta si desea conservar los dibujos (se convierten en elementos de tinta en las diapositivas) o descartarlos.

## Ensayar con intervalos

1. Haga clic en **Presentación con diapositivas > Ensayar intervalos**. La presentación comienza con un pequeño panel de temporizador que muestra el **Tiempo de diapositiva** y el **Tiempo total**; un botón **Pausa** detiene el reloj mientras hace un descanso.
2. Practique su charla, avanzando como lo haría normalmente. Se registra el tiempo que dedica a cada diapositiva.
3. Pulse **Escape** (o avance más allá de la última diapositiva). Un resumen de **Horarios de ensayo** lista cada diapositiva con su tiempo y su total.
4. Haga clic en **Guardar horarios** para almacenar los tiempos por diapositiva (usados para el avance automático cuando **Usar tiempos, si están presentes** está activado), o en **Desechar**.

## Presentar a espectadores remotos

**Presentación con diapositivas > Difundir** (también **Presente en línea** en el menú Presentar) inicia una sesión unidireccional en la que los espectadores siguen sus diapositivas en vivo desde un enlace. Consulte [Colaboración](/es/user/collaboration#difundir-una-presentacion).

## Salir

Pulse **Escape** o haga clic en la **X** de la barra de herramientas de la presentación. Si tiene anotaciones, aparece el cuadro de diálogo para conservarlas o descartarlas; si estaba ensayando, aparece el resumen de intervalos.

## A continuación

- Guarde un documento para el público o una grabación de su presentación: [Exportación](/es/user/exporting)
- Lista completa de teclas: [Atajos de teclado](/es/user/shortcuts)
