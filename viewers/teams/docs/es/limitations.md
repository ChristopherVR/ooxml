# Limitaciones

OpenTeams es una implementación temprana. Esta página enumera lo que no hace, para que puedas decidir antes de depender de él.

## No es Microsoft Teams

OpenTeams no es Microsoft Teams, no está afiliado a Microsoft y no afirma tener paridad de funciones. Implementa:

- canales, publicaciones con respuestas y paneles de hilos, reacciones, edición y borrado solo por parte del autor, búsqueda y contadores de no leídos;
- chats 1:1 y de grupo, cada uno en su propia sala a la que solo pueden unirse sus miembros;
- un feed de actividad y notificaciones dentro de la aplicación (aviso, notificación del navegador con la pestaña abierta, sonido);
- un calendario con reuniones, repeticiones, recordatorios y exportación `.ics`;
- presencia: quién está en línea, quién escribe y una disponibilidad que tú fijas (disponible, ocupado, ausente);
- reuniones en un canal: audio, vídeo, pantalla compartida y levantar la mano, con una pantalla previa a la unión;
- archivos compartidos en la conversación como enlaces, con los archivos de Word, Excel, PowerPoint y Visio reconocidos por nombre.

**No** implementa grabación, desenfoque de fondo, subtítulos en directo, notificaciones push ni por correo (las notificaciones funcionan solo con la aplicación abierta), invitaciones fuera del espacio de trabajo ni confirmaciones de asistencia (RSVP).

## Espacio de contenido

Word se previsualiza en solo lectura; Excel ofrece edición local y guardado de nuevas copias en el canal cuando hay almacenamiento configurado; Visio usa su visor existente. No hay coedición, sobrescritura ni guardado con control de versiones. PowerPoint necesita una página de visor incrustado configurada por el anfitrión. Las integraciones que faltan se indican en el panel. Las lecturas de texto se limitan a 2 MiB y las lecturas nativas de Office a 32 MiB. La interfaz está disponible en inglés, alemán, español, francés y chino simplificado (propiedad `locale`); el texto escrito por las personas nunca se traduce.

Las vistas de archivos admiten subida directa, creación de Excel en blanco y búsqueda de metadatos cuando hay almacenamiento configurado. Las subidas muestran un estado de ocupado y permiten reintentar, sin progreso en bytes ni cancelación. Los lotes fallidos pueden dejar objetos de almacenamiento sin referencia; el servidor de referencia no tiene un contrato de limpieza para esos objetos. No están implementados las carpetas, la indexación del contenido de los documentos, el historial de versiones de archivos ni los permisos de archivo autenticados.

Markdown admite encabezados, viñetas, citas, código delimitado, énfasis básico, enlaces web, listas de tareas planas, tablas de nivel superior y enlaces web relativos al archivo. No implementa CommonMark/GFM completo, estructuras anidadas ni enlaces de referencia. El HTML sin procesar se muestra como texto.

Los sitios web y los adjuntos HTML usan un sandbox sin permiso de mismo origen, ventanas emergentes ni navegación del nivel superior. Algunos sitios bloquean la incrustación o requieren permisos que este sandbox no concede. La apertura externa sigue disponible. Las definiciones de pestañas de archivos y sitios web pueden compartirse en los canales; la plataforma de aplicaciones incrustadas, la autenticación y los permisos de Microsoft Teams no están implementados. Las comprobaciones del creador de una pestaña se ejecutan en el cliente y requieren autorización en el servidor para poder hacerse cumplir.

## Seguridad y privacidad

- **Sin cifrado de extremo a extremo.** El chat viaja como actualizaciones de Yjs a través de tu servidor, que puede leerlas y las almacena. Los archivos subidos se almacenan tal cual. El contenido multimedia de las llamadas va por DTLS-SRTP entre pares (un relé TURN no puede descifrarlo).
- **La autorización es tarea del servidor.** Las comprobaciones de autor y de rol se ejecutan en el cliente; un cliente hostil puede escribir cualquier cosa que el servidor acepte.
- **Las cuentas son opcionales.** El servidor de referencia puede firmar sesiones para usuarios configurados (o tu función `authenticate`) y tomar el id de usuario de la sesión, y `authorizeSync` puede decidir el acceso por sala. Sin ellas, su token es un único secreto compartido almacenado en el navegador, no una autenticación por usuario, y los ids de usuario son declaraciones. No hay SSO, restablecimiento de contraseñas ni control de acceso por equipo y canal. Consulta [lo que el servidor de referencia no hace](/es/server#lo-que-el-servidor-de-referencia-no-hace).
- **La presencia es orientativa y no está autenticada**, como cualquier awareness de Yjs: un cliente puede declarar cualquier nombre.
- Todo lo que envía un par se vuelve a validar al leerlo y se muestra como texto, nunca como marcado.

## Escala

- **Las llamadas en malla** conectan a cada participante con todos los demás. No necesita servidor multimedia y funciona para un puñado de personas (el límite por defecto es 12), no para un webinar.
- **El historial crece.** El servidor de referencia mantiene una instantánea por sala y nunca la compacta.
- **Un solo proceso de servidor.** El servidor de referencia mantiene las salas en memoria y no escala horizontalmente.

## Modo local y las demos en vivo

`mode: 'local'` usa `BroadcastChannel`, de modo que solo se ven entre sí las pestañas del mismo navegador y del mismo sitio. No se envía nada a ningún servidor, los archivos se comparten solo por nombre (no hay dónde subirlos) y el documento compartido se guarda en el `localStorage` de ese navegador. Las [demos en vivo](/demos) funcionan así porque GitHub Pages no puede alojar un servidor.

## Llamadas entre redes

El STUN público funciona en redes abiertas. Detrás de un NAT simétrico o de cortafuegos estrictos necesitas tu propio servidor TURN; sin él, algunas llamadas no se conectarán. Las cámaras y los micrófonos requieren `localhost` o HTTPS.
