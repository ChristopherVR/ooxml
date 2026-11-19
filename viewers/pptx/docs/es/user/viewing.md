---
title: Visualizar presentaciones
description: Navegue por las diapositivas con las miniaturas, el teclado y el clasificador de diapositivas; haga zoom y ajuste el lienzo; entienda las diapositivas ocultas y lea las notas del orador.
---

# Visualizar presentaciones

Esta página trata sobre la lectura de una presentación: moverse entre diapositivas, hacer zoom, diapositivas ocultas y notas del orador. Para saber cómo abrir un archivo en primer lugar, consulte [Abrir un archivo](/es/user/#abrir-un-archivo).

Una vez cargado un archivo, la diapositiva activa se renderiza en el **lienzo de diapositivas** del centro, con una tira de miniaturas en el **panel de diapositivas** a la izquierda.

![Slide canvas showing a loaded presentation](/user-guide/viewing-slide-canvas.jpg)

Los gráficos, tablas, formas, SmartArt y medios se renderizan como contenido en vivo, así que puede seleccionar el texto de la diapositiva y copiarlo como en cualquier página web.

![A chart rendered on a slide](/docs-shots/chart-slide.jpg)

## Navegar por las diapositivas

Muévase entre diapositivas de cualquiera de estas formas:

- **Panel de diapositivas** - haga clic en cualquier miniatura de la izquierda para saltar directamente a esa diapositiva.
- **Teclas de flecha** - con nada seleccionado en el lienzo, pulse **Flecha izquierda** / **Flecha derecha** para pasar a la diapositiva anterior / siguiente.
- **Barra de estado** - el contador de la parte inferior izquierda indica **Diapositiva x de y** para que siempre sepa dónde está.

![Navigating between slides](/user-guide/viewing-navigation.gif)

El panel de diapositivas resalta la diapositiva actual con una barra de color en su borde izquierdo.

::: tip Las teclas de flecha tienen doble función
En modo de edición las teclas de flecha desplazan el elemento seleccionado. Solo cambian de diapositiva cuando no hay nada seleccionado. Pulse primero **Escape** para anular la selección.
:::

## Zoom y ajuste

Los controles de zoom están en el extremo derecho de la **barra de estado**:

| Control                  | Qué hace                                                                                               |
| ------------------------ | ------------------------------------------------------------------------------------------------------ |
| **Alejar** (botón menos) | Reduce el zoom por pasos.                                                                              |
| **Botón de porcentaje**  | Muestra el nivel de zoom actual; haga clic para **ajustar a la ventana** el área de lienzo disponible. |
| **Acercar** (botón más)  | Aumenta el zoom por pasos.                                                                             |

![Zoom controls in the status bar](/user-guide/viewing-status-bar.jpg)

También puede:

- Mantener **Ctrl** (Cmd en macOS) y girar la rueda del ratón sobre el lienzo para hacer zoom con suavidad alrededor del puntero.
- Usar **pestaña Vista > Ajustar a la ventana** en la cinta.
- Pellizcar para hacer zoom en dispositivos táctiles.

Para moverse por una diapositiva más grande que el lienzo, desplácese con la rueda, arrastre con el botón central del ratón o mantenga pulsada la tecla **Espacio** y arrastre. En una vista de solo lectura, donde no hay nada que seleccionar, también puede arrastrar la diapositiva con el botón izquierdo del ratón.

El zoom va del 10 % al 500 %, donde el 100 % ajusta la diapositiva al lienzo. Al ir a otra diapositiva se mantiene el zoom y la nueva diapositiva se muestra desde su esquina superior izquierda.

## Modos de vista

La barra de estado tiene tres botones de vista:

- **Vista normal** - el diseño estándar con el panel de diapositivas, el lienzo y el inspector.
- **Clasificador de diapositivas** - una cuadrícula a pantalla completa con todas las diapositivas.
- **Presentación con diapositivas** - inicia una presentación a pantalla completa (consulte [Presentar](/es/user/presenting)).

### El clasificador de diapositivas

El clasificador de diapositivas es la mejor manera de obtener una visión general de una presentación grande:

1. Haga clic en **Clasificador de diapositivas** en la barra de estado.
2. Explore la cuadrícula. Cuando la presentación tiene secciones, las diapositivas se agrupan bajo los encabezados de sus secciones.
3. **Haga doble clic** en una diapositiva para abrirla en la vista normal, o pulse **Escape** para cerrar el clasificador.

Dentro del clasificador también puede:

- Arrastrar diapositivas para reordenarlas (modo de edición).
- Hacer clic para seleccionar, **Mayús + clic** para seleccionar un intervalo y clic con el botón derecho para las acciones de diapositiva (duplicar, ocultar, eliminar; consulte [Gestionar diapositivas](/es/user/editing#gestionar-diapositivas)).
- Ampliar las miniaturas con el control deslizante de la barra de herramientas inferior, o pulsar **Ctrl +** / **Ctrl -**.

## Diapositivas ocultas

Las diapositivas marcadas como ocultas permanecen en el archivo pero se omiten durante una presentación.

- En el panel de diapositivas y en el clasificador aparecen atenuadas, con un patrón de rayado diagonal y un icono de ojo tachado.
- Aun así puede abrir, leer y editar una diapositiva oculta con normalidad.
- Para ocultar o mostrar diapositivas, haga clic con el botón derecho en ellas en el clasificador de diapositivas y elija **Ocultar diapositivas** o **Mostrar diapositivas** (modo de edición).

## Leer las notas del orador

1. Haga clic en el botón **Notas** de la barra de estado para abrir el panel de notas bajo el lienzo.
2. Las notas del orador de la diapositiva actual aparecen allí; desplácese si son largas.
3. Haga clic de nuevo en **Notas** para cerrar el panel.

Durante una presentación, las notas se le muestran a usted (no al público) en la **vista del presentador**; consulte [Presentar](/es/user/presenting#vista-del-presentador).

## Imágenes y elementos multimedia vinculados {#linked-pictures-and-media}

Algunas presentaciones vinculan imágenes, audio o vídeo a una dirección web en lugar de guardarlos en el archivo. Como en Microsoft PowerPoint, no se descargan al abrir la presentación, porque la descarga le indica a ese servidor que la ha abierto. Una barra bajo la cinta avisa de que se han bloqueado las referencias a imágenes y otros elementos multimedia externos. Aparece aunque el visor no tenga barra de herramientas.

- Haga clic en **Habilitar contenido** para descargarlos en esta presentación. No se vuelve a abrir nada: sus cambios, guardados o no, la diapositiva en pantalla y el historial de deshacer se quedan como están. En una sesión compartida, cada persona decide por sí misma.
- Haga clic en el botón de cierre para mantenerlos bloqueados.
- La elección dura hasta que abre otra presentación. Para descargar el contenido vinculado sin preguntar, active **Archivo > Opciones > Centro de confianza > Permitir contenido externo**.

## A continuación

- Haga cambios en las diapositivas: [Editar diapositivas](/es/user/editing)
- Ejecute una presentación: [Presentar](/es/user/presenting)
