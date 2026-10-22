---
title: Guía de usuario
description: Una orientación sobre el editor pptx-viewer - las partes de la pantalla, cómo abrir un archivo y dónde encontrar cada función.
---

# Guía de usuario

Esta guía es para personas que **usan** la aplicación pptx-viewer para trabajar con presentaciones de PowerPoint (`.pptx`) en el navegador. Cubre la visualización, la edición, la presentación, la exportación y la colaboración. No se necesita escribir código.

Si es desarrollador y está integrando el visualizador en su propia aplicación, consulte en su lugar la [documentación para desarrolladores del visualizador React](/react/).

::: info Las funciones dependen de cómo se configure la aplicación
pptx-viewer es un componente configurable, por lo que los botones y paneles exactos que ve dependen de cómo el desarrollador que creó su aplicación activó las funciones. La edición solo está disponible cuando la aplicación la habilita (aparece una insignia **Solo lectura** en la barra de herramientas cuando no lo está), y los controles de colaboración solo aparecen cuando se habilita el uso compartido. Cuando una función es opcional, esta guía lo indica.
:::

## Qué es el editor

pptx-viewer abre archivos de PowerPoint directamente en su navegador y renderiza cada diapositiva como HTML y SVG reales, no como una imagen plana. El texto se mantiene nítido a cualquier zoom, se puede seleccionar y copiar, y funciona con lectores de pantalla. En modo de edición puede cambiar las diapositivas y guardar el resultado de nuevo en un archivo `.pptx` estándar.

![The editor with a presentation open](/docs-shots/editor.jpg)

## Abrir un archivo

1. Arrastre un archivo `.pptx` al área de colocación etiquetada **"Suelte un archivo .pptx aquí o haga clic para examinar"**, o haga clic en el área para abrir el selector de archivos de su sistema.
2. Espere un momento mientras se analiza el archivo. La primera diapositiva aparece en el lienzo con las miniaturas a la izquierda.

También puede hacer clic en **"o crear una nueva presentación"** para empezar desde una presentación en blanco de una sola diapositiva.

::: tip Su archivo permanece en su equipo
El archivo se procesa íntegramente en el navegador. No se sube nada a un servidor: abrir, editar y guardar ocurren localmente.
:::

::: warning Archivos admitidos
La aplicación abre `.pptx` y sus parientes (`.ppsx`, `.pptm`, `.potx`), además de **`.ppt` binario heredado** de PowerPoint 97-2003, que se convierte al cargarlo. Guardar usa `.pptx` por defecto sea cual sea el archivo que abrió (una presentación abierta como `report.ppt` se ofrece de vuelta como `report.pptx`), pero Guardar como también ofrece **Presentación de PowerPoint 97-2003 (`.ppt`)** como formato explícito junto a `.ppsx` y `.pptm`; consulte [Limitaciones](/es/guide/limitations) para ver las carencias de fidelidad restantes en ese formato (acciones de clic en hipervínculos y elementos sin equivalente binario).

Los archivos `.pptx` protegidos con contraseña se abren una vez que proporciona la contraseña: un cuadro de diálogo titulado **Archivo cifrado** la solicita. Los archivos **`.ppt`** protegidos con contraseña no se pueden abrir en absoluto, porque el esquema de cifrado antiguo no se descifra.
:::

Si la aplicación encuentra trabajo sin guardar de una visita anterior, aparece un aviso titulado **Cambios sin guardar recuperados** con los botones **Restaurar** y **Descartar**. Consulte [Autoguardado y recuperación](/es/user/editing#autoguardado-y-recuperacion).

## Un recorrido por la pantalla

![Full editor layout](/user-guide/overview-full-layout.jpg)

De arriba abajo:

### Barra de título

La franja superior muestra el nombre del archivo y su estado de guardado (por ejemplo "Guardado en este PC" o "Cambios sin guardar"). Cuando la edición está habilitada, también contiene:

- El interruptor de **Autoguardado** (Activado/Desactivado).
- Los botones rápidos **Guardar**, **Deshacer** y **Rehacer**.
- Un cuadro de búsqueda, **"Dime qué quieres hacer"**. Escriba el nombre de un comando (por ejemplo "tabla") para obtener una lista en la que se puede hacer clic con las acciones coincidentes, o pulse **Intro** para buscar el texto dentro de sus diapositivas.

### Cinta (barra de herramientas)

Una cinta con pestañas bajo la barra de título. Las pestañas, en orden:

**Archivo, Inicio, Insertar, Dibujar, Diseño, Transiciones, Animaciones, Presentación con diapositivas, Grabar, Revisar, Vista, Ayuda**

![Toolbar close-up](/user-guide/overview-toolbar.jpg)

- **Archivo** abre una vista Backstage a pantalla completa con las páginas Abrir, Guardar, Guardar como, Imprimir, Compartir y Exportar. Consulte [Exportación](/es/user/exporting).
- **Inicio** contiene las herramientas cotidianas: portapapeles, diapositivas y diseños nuevos, fuentes, formato de párrafo, buscar y reemplazar, dibujo y controles de organización. Consulte [Edición](/es/user/editing).
- **Insertar** añade elementos nuevos: cuadros de texto, formas, imágenes, medios, tablas, gráficos, SmartArt, ecuaciones y más.
- **Presentación con diapositivas** inicia una presentación. Consulte [Presentación](/es/user/presenting).

La cinta solo aparece en modo de edición. En el extremo derecho de la fila de pestañas están los botones **Grabar** y **Compartir**; la fila superior contiene los interruptores del panel de diapositivas, los comentarios y el inspector, el botón **Presentar**, un engranaje de Configuración y un menú de desbordamiento **"..."** que repite las acciones de exportación y utilidades.

### Panel de diapositivas (izquierda)

Una tira desplazable de miniaturas de diapositivas.

![Slides panel](/user-guide/overview-slides-panel.jpg)

- Haga clic en una miniatura para saltar a esa diapositiva.
- Arrastre las miniaturas para reordenarlas (modo de edición).
- Las diapositivas ocultas aparecen atenuadas, con un rayado diagonal y un icono de ojo tachado.
- El botón **Agregar diapositiva** de la parte inferior inserta una diapositiva nueva.
- Las diapositivas pueden agruparse en **secciones** con nombre y sus propios encabezados.

### Lienzo de diapositivas (centro)

El área grande donde se renderiza la diapositiva activa. Aquí se hace zoom y se desplaza la vista, y en modo de edición se seleccionan, mueven y editan elementos directamente. Las reglas, una cuadrícula y guías arrastrables se pueden activar desde la pestaña **Vista**.

### Inspector (derecha)

Un panel lateral sensible al contexto con tres pestañas:

- **Elementos** - una lista de capas de todos los elementos de la diapositiva, en orden de apilamiento. Haga clic en una fila para seleccionar ese elemento en el lienzo.
- **Propiedades** - ajustes de la selección actual. Sin nada seleccionado muestra los ajustes de la presentación, el tema, el tamaño de diapositiva, la transición de diapositiva y el fondo; con un elemento seleccionado muestra la posición y el tamaño más paneles específicos del tipo de elemento (texto, relleno y trazo, tabla, datos del gráfico, imagen, medios, etc.) y un panel de **Animación**.
- **Comentarios** - los hilos de comentarios asociados a la diapositiva actual, con controles para responder y resolver.

Active o desactive el inspector con el botón de panel de la fila superior de la barra de herramientas.

### Panel de notas (debajo del lienzo)

Un editor para las notas del orador de la diapositiva actual. Actívelo con el botón **Notas** de la barra de estado.

### Barra de estado (abajo)

De izquierda a derecha:

- El contador **Diapositiva x de y**.
- El estado de guardado: **Cambios sin guardar**, **Todo guardado**, **Guardando...** o **Guardado [hora]**.
- El interruptor **Notas**.
- Tres botones de vista: **Vista normal**, **Clasificador de diapositivas**, **Presentación con diapositivas**.
- El indicador de conexión cuando hay una sesión de colaboración activa.
- Controles de zoom: **Alejar**, un botón de porcentaje que también actúa como **Ajustar a la ventana**, y **Acercar**.

## Escritorio y móvil

::: warning Diseñado para escritorio
La cinta, el inspector y los cuadros de diálogo están diseñados para pantallas de tamaño de escritorio. En un teléfono o una tableta pequeña la aplicación cambia a una barra de herramientas compacta y hojas inferiores, y admite gestos táctiles (arrastrar, pellizcar para hacer zoom), pero la edición seria es más cómoda en una pantalla más grande.
:::

<div style="display: flex; gap: 1rem; align-items: flex-start;">
  <figure style="flex: 1;">
    <img src="/user-guide/overview-mobile-layout.jpg" alt="Mobile viewer layout" style="max-width: 280px;" />
    <figcaption>Visualizador móvil</figcaption>
  </figure>
  <figure style="flex: 1;">
    <img src="/user-guide/overview-mobile-inspector.jpg" alt="Mobile inspector bottom sheet" style="max-width: 280px;" />
    <figcaption>Inspector móvil (hoja inferior)</figcaption>
  </figure>
</div>

## A dónde ir a continuación

- [Visualizar presentaciones](/es/user/viewing) - navegue por las diapositivas, haga zoom, diapositivas ocultas y lectura de notas.
- [Editar diapositivas](/es/user/editing) - seleccione y formatee elementos, inserte contenido, gestione diapositivas, deshaga y autoguarde.
- [Presentar](/es/user/presenting) - ejecute una presentación, la vista del presentador y las anotaciones en vivo.
- [Exportación](/es/user/exporting) - guarde de nuevo en `.pptx`, exporte PDF, imágenes, GIF y video, e imprima.
- [Colaboración](/es/user/collaboration) - coedite en tiempo real y difunda una presentación.
- [Atajos de teclado](/es/user/shortcuts) - la referencia completa de atajos.
