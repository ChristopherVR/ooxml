import type { LandingCopy } from './types';

export const es: LandingCopy = {
	hero: {
		titleTop: 'Chat y llamadas de equipo,',
		titleAccent: 'en un servidor que tú ejecutas.',
		sub: 'OpenTeams es un espacio de trabajo de equipo para la web: canales y chat, presencia, reuniones WebRTC y archivos de Office compartidos en la conversación. Un componente web, bindings ligeros para seis frameworks y un pequeño servidor de referencia. Una implementación temprana: no es Microsoft Teams y no tiene cifrado de extremo a extremo.',
		start: { text: 'Empezar', href: '/es/getting-started' },
		demo: { text: 'Probar la demo', href: '/demo/' },
		frameLabel: '<teams-app> en modo local',
		tryLabel: 'Abrir la demo',
		imageAlt:
			'La demo de OpenTeams: una lista de canales, una conversación con respuestas y reacciones, y el cuadro de redacción',
	},
	liveDemo: {
		title: 'Pruébalo aquí mismo',
		note: 'Demo en vivo',
		introBefore:
			'Esta es la aplicación real ejecutándose en tu navegador, sin ningún servidor detrás: el chat, la presencia y las llamadas viajan solo entre marcos y pestañas de este navegador (BroadcastChannel). Cambia a dos personas para poner a Ada y a Bob en una misma sala, cada uno en el framework que elijas (por ejemplo React para Ada y Vue para Bob), y observa cómo un mensaje pasa de un panel al otro. El uso real necesita un servidor; consulta ',
		introLink: { text: 'aporta tu propio servidor', href: '/es/server' },
		introAfter: '.',
		frameworkLabel: 'Framework de Ada',
		soloTab: 'Una persona',
		duoTab: 'Dos personas',
		guestPicker: 'Framework de Bob',
		load: 'Cargar la demo en vivo',
		openFull: 'Abrir la aplicación completa',
		frameTitle: 'Demo de OpenTeams en {framework}',
		hintDuo:
			'Dos aplicaciones independientes, cada una construida con su propio binding de framework, se unen a una sala por su nombre y la comparten mediante BroadcastChannel. El framework no importa: hablan el mismo protocolo del núcleo. Este es el modo local del núcleo, no un servicio alojado: con tu propio servidor, las mismas dos personas pueden estar en máquinas distintas.',
		hintSolo:
			'Todo se ejecuta en el cliente y permanece en este navegador. Abre la aplicación completa o cambia a dos personas para ver la presencia y el chat entre dos usuarios.',
	},
	embed: {
		title: 'Dos formas de incrustarlo',
		note: 'Aplicaciones de demostración',
		openDemo: 'Abrir la demo',
		demos: [
			{
				title: 'Vanilla: la aplicación completa',
				copy: 'Un único elemento <teams-app>: canales, publicaciones con respuestas y reacciones, archivos, búsqueda, presencia y reuniones. Ábrelo en dos pestañas para ver a dos personas conversar.',
				href: '/demo/',
				link: '',
			},
			{
				title: 'React: componente y hook',
				copy: 'El componente <Teams /> a la izquierda y un panel escrito a mano sobre el hook useTeams() a la derecha, ambos sobre un mismo espacio de trabajo.',
				href: '/demo-react/',
				link: '',
			},
		],
	},
	features: {
		title: 'Un espacio de trabajo que puedes alojar',
		note: 'Qué hace',
		items: [
			{
				title: 'Canales y chat',
				copy: 'Canales, publicaciones con respuestas, reacciones, edición y borrado solo por el autor, búsqueda y contadores de no leídos, todo en un único documento de Yjs que fusiona las ediciones concurrentes.',
				href: '/es/getting-started',
				link: 'Primeros pasos',
				icon: 'chat',
			},
			{
				title: 'Presencia',
				copy: 'Quién está en línea, indicadores de escritura y una disponibilidad que fijas tú mismo (disponible, ocupado, ausente). La presencia es orientativa, como cualquier awareness de Yjs.',
				href: '/architecture',
				link: 'Arquitectura',
				icon: 'people',
			},
			{
				title: 'Reuniones sobre WebRTC',
				copy: 'Audio, vídeo, pantalla compartida y levantar la mano en una llamada en malla completa: sin servidor multimedia, razonable hasta una docena de personas. El TURN lo ejecutas tú.',
				href: '/es/server#llamadas-stun-y-turn',
				link: 'Llamadas y TURN',
				icon: 'video',
			},
			{
				title: 'Archivos en la conversación',
				copy: 'Los adjuntos se suben a tu servidor (o a tu propio cargador) y se publican como enlaces. Los archivos de Word, Excel, PowerPoint y Visio se reconocen para que un anfitrión pueda abrirlos.',
				href: '/es/server#archivos',
				link: 'Almacenamiento de archivos',
				icon: 'file',
			},
			{
				title: 'Aporta tu propio servidor',
				copy: 'Tres contratos pequeños: sincronización y-websocket, un relé de señalización JSON y almacenamiento de archivos opcional. openteams-server cumple los tres en un único proceso de Node.',
				href: '/es/server',
				link: 'Guía del servidor',
				icon: 'server',
			},
			{
				title: 'Seis frameworks',
				copy: 'Bindings de React, Vue, Angular, Svelte, Solid y vanilla. Cada uno te da el componente completo y un hook sin procesar sobre el mismo almacén del cliente.',
				href: '/frameworks/react',
				link: 'Guías por framework',
				icon: 'code',
			},
		],
	},
	limits: {
		title: 'Lo que no es',
		note: 'Límites honestos',
		items: [
			'Sin cifrado de extremo a extremo: tu servidor almacena y puede leer todos los mensajes y archivos.',
			'La autorización es tarea del servidor; el servidor de referencia tiene un único token compartido, no cuentas de usuario.',
			'Las llamadas en malla sirven para un puñado de personas, no para un webinar.',
			'Sin chats 1:1, notificaciones, calendario, grabación, desenfoque de fondo ni subtítulos en directo.',
			'No es Microsoft Teams, no está afiliado a Microsoft y no afirma tener paridad de funciones.',
		],
		all: { text: 'Todas las limitaciones', href: '/es/limitations' },
	},
	quickstart: {
		title: 'Un espacio de trabajo en unas pocas líneas',
		note: 'Inicio rápido',
		introBefore:
			'Instala el binding de tu framework y renderiza el componente. El modo local no necesita servidor; para un uso real apunta ',
		introAfter:
			' a tus propios endpoints de sincronización y señalización (el servidor de referencia ofrece ambos). Angular, Svelte y Solid tienen las mismas opciones.',
		fullGuide: 'Guía completa: {framework}',
	},
	faq: {
		title: 'Preguntas frecuentes',
		note: 'FAQ',
		items: [
			{
				q: '¿Es gratis para uso comercial?',
				a: 'El repositorio tiene licencia Apache-2.0. Los paquetes se publican en npm como openteams-*-viewer y openteams-server.',
			},
			{
				q: '¿Necesita un servidor?',
				a: 'El uso real sí: aportas tu propio servidor para la sincronización (y-websocket), la señalización y el almacenamiento de archivos opcional. El servidor de referencia, openteams-server, cumple los tres en un único proceso de Node. Sin servidor, las demos funcionan en un modo local que solo conecta pestañas de un mismo navegador.',
				href: '/es/server',
				link: 'Aporta tu propio servidor',
			},
			{
				q: '¿Tiene cifrado de extremo a extremo?',
				a: 'No. El chat viaja como actualizaciones de Yjs a través de tu servidor, que puede leerlas y almacenarlas. El contenido multimedia de las llamadas va por DTLS-SRTP entre pares.',
				href: '/es/limitations',
				link: 'Limitaciones',
			},
			{
				q: '¿Es esto Microsoft Teams?',
				a: 'No. OpenTeams no es Microsoft Teams, no está afiliado a Microsoft y no afirma tener paridad de funciones. Tiene canales, publicaciones con respuestas, reacciones, archivos, búsqueda, presencia y reuniones con pantalla compartida.',
			},
			{
				q: '¿Cuántas personas pueden unirse a una llamada?',
				a: 'Las llamadas son una malla completa sin servidor multimedia, así que sirven para un puñado de personas (límite por defecto de 12), no para un webinar. El TURN lo ejecutas tú.',
				href: '/es/server#llamadas-stun-y-turn',
				link: 'Llamadas y TURN',
			},
			{
				q: '¿Qué frameworks son compatibles?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS y JavaScript puro. Comparten un único elemento <teams-app> y un único cliente del núcleo; cada binding expone además un hook sin procesar para que construyas tu propia interfaz.',
				href: '/frameworks/react',
				link: 'Guías por framework',
			},
			{
				q: '¿Funciona con archivos de Word, Excel, PowerPoint y Visio?',
				a: 'Los archivos se comparten como enlaces. Los archivos de Office se reconocen por su nombre para que una aplicación anfitriona pueda abrirlos en el visor correspondiente; OpenTeams no los edita por sí mismo.',
			},
		],
	},
	finale: {
		title: 'Tu espacio de trabajo, tu servidor.',
		introBefore: 'Instala un binding, arranca el servidor de referencia con ',
		introAfter:
			' y prueba la demo con dos pestañas. Con licencia Apache-2.0, TypeScript estricto y honesto sobre las carencias.',
		start: { text: 'Empezar', href: '/es/getting-started' },
		github: 'Ver en GitHub',
		columns: [
			{
				title: 'Producto',
				links: [
					{ text: 'Demo en vivo', href: '/demo/', external: false },
					{ text: 'Aporta tu propio servidor', href: '/es/server', external: false },
					{ text: 'Colaboración', href: '/es/collaboration', external: false },
					{ text: 'Versiones', href: '/releasing', external: false },
				],
			},
			{
				title: 'Documentación',
				links: [
					{ text: 'Primeros pasos', href: '/es/getting-started', external: false },
					{ text: 'Arquitectura', href: '/architecture', external: false },
					{ text: 'API del elemento', href: '/es/api', external: false },
					{ text: 'Limitaciones', href: '/es/limitations', external: false },
				],
			},
			{
				title: 'Comunidad',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams',
						external: true,
					},
					{
						text: 'Incidencias',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Licencia',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/LICENSE',
						external: true,
					},
				],
			},
		],
		license: '© 2026 ChristopherVR · Apache-2.0',
		disclaimer: 'OpenTeams · no es Microsoft Teams, no está afiliado a Microsoft',
	},
};
