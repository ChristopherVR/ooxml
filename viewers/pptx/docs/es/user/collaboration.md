---
title: Colaboración
description: Inicie o únase a una sesión de coedición en vivo, interprete los indicadores de presencia, entienda cómo se combinan las ediciones simultáneas y difunda una presentación.
---

# Colaboración

pptx-viewer admite la coedición en tiempo real: varias personas abren la misma presentación y ven los cambios de las demás a medida que ocurren.

::: info Opcional y configurado por el desarrollador
La colaboración es una función opcional. Si no ve un botón **Compartir**, la aplicación que usa se creó sin ella. Desarrolladores: consulte [/react/collaboration](/react/collaboration).
:::

## Iniciar una sesión

1. Haga clic en el botón **Compartir** en el extremo derecho de la fila de pestañas de la cinta (o en **Archivo > Compartir > Compartir con personas**). El cuadro de diálogo **Compartir presentación** se abre en su pestaña **Crear sesión**.
2. Rellene:
   - **Nombre de la sesión** - el identificador de la sala (letras, números, guiones y guiones bajos). Los colaboradores necesitan este nombre.
   - **Su nombre para mostrar** - cómo aparece ante los demás.
   - **Servidor de colaboración** - opcional. Déjelo vacío para una sesión entre pares, o introduzca la dirección del servidor de colaboración de su organización (una URL `wss://...`).
3. Haga clic en **Iniciar uso compartido**.

![Share/collaboration dialog](/user-guide/collaboration-share-dialog.jpg)

Cuando la sesión está activa, el cuadro de diálogo muestra un **Enlace para compartir** con un botón **Copiar URL**. Envíe ese enlace a sus colaboradores. El cuadro de diálogo también lista los **Usuarios conectados** y en qué diapositiva está cada persona. El botón Compartir de la barra de herramientas se vuelve verde y muestra el número de participantes mientras el uso compartido está activo.

Para terminar la sesión, abra de nuevo el cuadro de diálogo y haga clic en **Detener uso compartido**.

## Unirse a una sesión

Abra un enlace para compartir que alguien le haya enviado, o únase manualmente:

1. Haga clic en **Compartir** y cambie a la pestaña **Unirse a una sesión**.
2. Pegue el enlace de invitación, o introduzca el ID de la sala, en **Enlace de invitación o ID de sala**.
3. Introduzca su nombre para mostrar y, si la sesión usa uno, la dirección del servidor de colaboración.
4. Haga clic en **Unirse a una sesión**.

## Indicadores de presencia

Mientras hay una sesión activa verá:

- **Barra de avatares** - círculos de colores en la barra de herramientas para cada participante (hasta cinco, y luego un chip de desbordamiento "+N"). Su propio avatar tiene un borde blanco.
- **Cursores remotos** - el puntero de cada colaborador se mueve por el lienzo con una etiqueta de nombre de color. Los cursores solo aparecen para las personas que ven la misma diapositiva que usted.
- **Selecciones remotas** - cuando alguien selecciona un elemento, este recibe un contorno de color con su nombre, para que sepa en qué está trabajando.
- **Barra de seguimiento** - una franja en la parte superior etiquetada **Seguir a un colaborador**. Haga clic en una persona para seguirla: su vista cambia de diapositiva cada vez que ella lo hace. Haga clic en **Detener** para volver a navegar por su cuenta.
- **Estado de la conexión** - un punto en la barra de estado: **Conectado** en verde (con el número de usuarios), **Conectando...** en amarillo, **Desconectado** en gris, o **Error de conexión** en rojo con un botón **Reintentar**.

## Cómo se comporta la edición simultánea

Las ediciones se combinan automáticamente con estructuras de datos libres de conflictos:

- Dos personas que cambian diapositivas distintas, o elementos distintos de la misma diapositiva, nunca se sobrescriben entre sí.
- Dos personas que escriben en el mismo cuadro de texto se combinan a nivel de carácter, como en un editor de documentos compartido.
- Si se interrumpe su conexión, puede seguir editando; sus cambios se sincronizan cuando la conexión vuelve. Observe el punto de estado: mientras muestre **Desconectado**, los demás aún no ven sus últimas ediciones.

::: tip Coordine los cambios grandes
Combinar le protege de perder ediciones, no de las sorpresas. Eliminar una diapositiva en la que otra persona está trabajando la elimina igualmente para todos, así que acuerden quién se encarga de cada sección en reestructuraciones importantes.
:::

## Entre pares o servidor de retransmisión

El campo **Servidor de colaboración** decide cómo se conectan los participantes:

- **Vacío (entre pares)** - las pestañas del navegador del mismo dispositivo siempre se encuentran entre sí, y las personas de otros dispositivos se conectan directamente con usted a través de servidores públicos de señalización WebRTC. Es la forma más rápida de probar la colaboración; la demo alojada funciona así.
- **Dirección de servidor definida** - todos retransmiten a través de un servidor de colaboración dedicado. Úselo cuando su organización tenga uno: funciona a través de cortafuegos estrictos y NAT donde fallan las conexiones directas entre pares, mantiene la sala disponible mientras haya algún participante conectado y la aplicación anfitriona puede combinarlo con autenticación y persistencia en el servidor.

::: warning Use wss:// en páginas seguras
Cuando la aplicación se sirve por `https`, la dirección del servidor debe usar `wss://` (un WebSocket seguro), o el navegador bloqueará la conexión.
:::

## Difundir una presentación

Difundir es una colaboración unidireccional: usted presenta, los espectadores siguen.

1. Haga clic en **Presentación con diapositivas > Difundir** (o **Presente en línea** en el menú Presentar). Se abre el cuadro de diálogo **Transmitir presentación**.
2. Compruebe el nombre de **Sesión de transmisión** rellenado de antemano y su **Nombre del presentador**, y defina o borre el **Servidor de colaboración** como antes.
3. Haga clic en **Iniciar transmisión**. La aplicación entra en modo de presentación y el cuadro de diálogo muestra un **Enlace del espectador** con **Copiar URL**.
4. Comparta el enlace de espectador. Quienes lo abran verán avanzar sus diapositivas en tiempo real; el cuadro de diálogo cuenta a sus espectadores.
5. Haga clic en **Detener transmisión** cuando termine.

## Los comentarios funcionan junto con la edición en vivo

Los comentarios (véase [Edición > Comentarios](/es/user/editing#comentarios)) se guardan con la presentación y son una buena forma de dejar comentarios a los colaboradores que no están en línea al mismo tiempo.

## A continuación

- [Editar diapositivas](/es/user/editing)
- [Atajos de teclado](/es/user/shortcuts)
