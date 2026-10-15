import type { LandingCopy } from './types';

export const es: LandingCopy = {
	hero: {
		kicker: 'Beta pública · Apache-2.0 · TypeScript',
		titleTop: 'Visualización de .vsdx,',
		titleAccent: 'lista para incrustar.',
		sub: 'Lleve los diagramas de Visio a su aplicación web. Un único componente web <visio-viewer> compartido y un único renderizador SVG para React, Vue, Angular, Svelte, Solid y JavaScript puro. Explore páginas, inspeccione formas y haga las ediciones de texto sin formato admitidas; los adaptadores de framework solo conectan el ciclo de vida y los eventos.',
		start: { text: 'Empezar', href: '/es/getting-started' },
		demo: 'Demostración en vivo',
		scroll: 'Desplazar',
		frameCaption: 'Flujo de trabajo de ejemplo · en vivo en el navegador',
		frameTry: 'Pruébelo',
		frameAlt:
			'El espacio de trabajo de visio-viewer con el ejemplo original, la navegación de páginas, la cinta de opciones y el inspector de formas',
		copyLabel: 'Copiar',
		copiedLabel: 'Copiado',
		entryLabel: 'instalar',
		notPublished:
			'Publicado en npm. Cada paquete de framework es autónomo: incluye el visor compartido, su adaptador y sus workers.',
	},
	features: {
		kicker: 'Funciones',
		title: 'Lo que existe hoy.',
		items: [
			{
				title: 'Seis enlaces nativos',
				copy: 'React, Vue, Angular, Svelte, Solid o JavaScript puro. Adaptadores de ciclo de vida ligeros mantienen toda la lógica del formato en ooxml-core y en el visor compartido.',
				link: { text: 'Guías de frameworks', href: '/frameworks/react' },
			},
			{
				title: 'Un diagrama en primer plano',
				copy: 'Cambie de página, encaje y amplíe, seleccione formas, busque texto visible, ajuste la visibilidad de las capas admitidas y use la cinta de opciones al estilo de Visio, la vista Backstage de Archivo, la ventana Formas y la ventana Panorámica y zoom.',
				link: { text: 'Demostraciones en vivo', href: '/demos' },
			},
			{
				title: 'Su navegador, sus archivos',
				copy: 'Abra un archivo .vsdx local y exporte la página actual como SVG. No hay subida de documentos, ni telemetría, ni carga de documentos remotos.',
				link: { text: 'Primeros pasos', href: '/es/getting-started' },
			},
			{
				title: 'Edición cuidadosa',
				copy: 'Reemplazo experimental de texto sin formato respaldado por el origen, ediciones de geometría acotadas, deshacer y rehacer acotados y copias VSDX originales o editadas de forma explícita. El contenido no admitido se rechaza, no se descarta.',
				link: { text: 'Límites de la edición', href: '/es/getting-started#editing' },
			},
			{
				title: 'Vista previa de .vsd heredado',
				copy: 'Los archivos binarios VSD v11 se cargan mediante el códec ole2 compartido de ooxml-core como vista previa de solo lectura, con estilos alternativos diagnosticados.',
				link: { text: 'Registro de capacidades', href: '/parity' },
			},
			{
				title: 'Uso compartido en vivo',
				copy: 'Archivo > Compartir inicia una sesión de Yjs sobre BroadcastChannel entre ventanas de un mismo navegador. Compartir entre dispositivos necesita un transporte de servidor aportado por el anfitrión, que no está construido.',
				link: { text: 'Colaboración', href: '/es/collaboration' },
			},
			{
				title: 'Exportación y preparación de impresión',
				copy: 'Exportación SVG de la página actual con rásteres incrustados y notas de compatibilidad visibles, además de instantáneas de página inmutables mediante la API. No hay cuadro de diálogo de impresión, ni exportación a PDF o mapa de bits.',
				link: { text: 'API del visor', href: '/es/api' },
			},
			{
				title: 'Limitaciones visibles',
				copy: 'Cada capacidad se registra con su evidencia y con lo que aún falta. Una cinta de opciones atractiva no se presenta como paridad con Visio.',
				link: { text: 'Registro de capacidades', href: '/parity' },
			},
		],
	},
	status: {
		kicker: 'Estado',
		title: 'Honestos sobre lo que falta.',
		copy: 'Esta es una beta pública, no paridad con Microsoft Visio. No están establecidos el dibujo general, el recálculo completo de ShapeSheet, el texto enriquecido, los grupos y la mayoría de las plantillas maestras, la exportación a PDF ni la fidelidad de reapertura en Visio nativo. El ejemplo generado demuestra interacciones; no prueba la fidelidad visual.',
		link: { text: 'Registro de capacidades', href: '/parity' },
		roadmapTitle: 'Dónde reside la lógica',
		roadmapCopy:
			'El análisis, la interpretación de ShapeSheet, la geometría y los comandos de edición residen en ooxml-core/visio. Este repositorio contiene solo la interfaz: un controlador, un renderizador SVG, un elemento y seis adaptadores ligeros.',
		roadmapLink: { text: 'Arquitectura', href: '/es/architecture' },
	},
	quickstart: {
		kicker: 'Primeros pasos',
		title: 'Monte un visor en unas pocas líneas.',
		copy: 'Cada adaptador reenvía las mismas propiedades, los mismos eventos y el mismo manejador imperativo del visor compartido. Dé una altura al contenedor y después pase un VisioDocument o cargue un File.',
		docsLabel: 'Guía del framework',
		buildTitle: 'Ejecútelo desde el código fuente',
		buildCopy:
			'Para ejecutar el área de pruebas o trabajar en el propio visor, clone el repositorio. Se requiere Node.js 22.12 o posterior.',
		buildCommands: 'npm ci --ignore-scripts\nnpm run dev',
	},
	demos: {
		kicker: 'Demostración en vivo',
		title: 'Pruébelo aquí mismo.',
		copy: 'Este es el visor real ejecutándose en su navegador: el área de pruebas creada a partir de este repositorio para cada adaptador de framework, incrustada en vivo. Abra el dibujo de ejemplo o un .vsdx local; no se sube nada.',
		frameworkLabel: 'Framework',
		soloTab: 'Visor',
		collabTab: 'Uso compartido',
		guestPicker: 'Framework de la ventana B',
		load: 'Cargar la demostración en vivo',
		loading: 'Cargando el visor en vivo',
		openFull: 'Abrir la aplicación completa',
		hostLabel: 'Ventana A',
		guestLabel: 'Ventana B',
		soloHint:
			'Use el dibujo de ejemplo o abra un archivo .vsdx propio. Al cambiar de framework se inicia un visor nuevo.',
		collabHint:
			'Dos ventanas de un mismo navegador, cada una en el framework que elija, ya unidas por nombre a una sesión de Archivo > Compartir. La ventana A abre el ejemplo y la ventana B lo recibe. Edite el texto de una forma en una y aparecerá en la otra. Usa un BroadcastChannel identificado por el nombre de la sesión; no hay servidor, solo funciona dentro de este navegador y no se sube nada.',
	},
	faq: {
		kicker: 'Preguntas frecuentes',
		title: 'Preguntas habituales.',
		items: [
			{
				q: '¿Es un reemplazo completo de Visio?',
				a: 'No. Es una implementación temprana. La geometría, el texto y las imágenes en caché que se admiten se renderizan por un único camino SVG. Los documentos complejos aún tienen carencias; el registro de capacidades separa la implementación, la evidencia y el trabajo pendiente.',
				link: { text: 'Registro de capacidades', href: '/parity' },
			},
			{
				q: '¿Puedo instalarlo desde npm?',
				a: 'Sí. En npm hay un paquete de visor autónomo por framework, por ejemplo npm install visio-react-viewer. Cada paquete tiene su propio control de versiones.',
				link: { text: 'Primeros pasos', href: '/es/getting-started' },
			},
			{
				q: '¿Abrir un archivo lo sube?',
				a: 'No. El visor lee en su navegador el File que usted selecciona. No hay subida de documentos, telemetría ni obtención de documentos remotos.',
			},
			{
				q: '¿Puedo editar y guardar un archivo .vsdx?',
				a: 'Hay disponible una edición experimental de texto sin formato y de geometría acotada para las formas admitidas respaldadas por el origen, con deshacer y rehacer acotados y descargas explícitas de copias VSDX. Se rechazan el texto vinculado a plantillas maestras, el texto enriquecido, los campos, los paquetes firmados y el contenido con macros. El archivo de origen nunca se sobrescribe y la reapertura en Visio nativo sigue sin verificarse.',
				link: { text: 'API del visor', href: '/es/api' },
			},
			{
				q: '¿Qué puedo exportar o imprimir?',
				a: 'La interfaz exporta SVG de la página actual y copias VSDX originales o editadas. La API puede preparar instantáneas de página inmutables para un futuro flujo de impresión; no abre un cuadro de diálogo de impresión. La exportación a PDF o a mapa de bits y la maquetación de impresora nativa no están implementadas.',
			},
			{
				q: '¿Pueden varias personas editar un mismo diagrama?',
				a: 'Solo de forma limitada: Archivo > Compartir enlaza ventanas del mismo navegador mediante una sesión de Yjs. Compartir entre dispositivos necesita un transporte de servidor aportado por el anfitrión que esta compilación no incluye.',
				link: { text: 'Colaboración', href: '/es/collaboration' },
			},
			{
				q: '¿Demuestra el ejemplo la fidelidad visual?',
				a: 'No. El ejemplo original, creado mediante programación, demuestra interacciones. Los fixtures generados no establecen la paridad visual con Microsoft Visio; compare los dibujos importantes con su representación en Visio nativo.',
			},
			{
				q: '¿Qué frameworks se admiten?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS y JavaScript puro. Comparten un único elemento de visor; cada adaptador solo gestiona el montaje, las actualizaciones de propiedades y el reenvío de eventos.',
				link: { text: 'Guías de frameworks', href: '/frameworks/react' },
			},
		],
	},
	finale: {
		kicker: 'Empiece',
		title: 'Entra un .vsdx. Sale un diagrama.',
		sub: 'Lea la guía, elija su framework y pruebe el área de pruebas con un dibujo propio. Con licencia Apache-2.0, TypeScript estricto y honestos sobre las carencias.',
		quick: { text: 'Leer la guía', href: '/es/getting-started' },
		github: 'Ver en GitHub',
		columns: [
			{
				title: 'Producto',
				links: [
					{
						text: 'Demostración en vivo',
						href: 'https://christophervr.github.io/ooxml/visio/demo/',
						external: true,
					},
					{ text: 'Demostraciones en vivo', href: '/demos' },
					{ text: 'Colaboración', href: '/es/collaboration' },
					{ text: 'Versiones', href: '/releasing' },
				],
			},
			{
				title: 'Documentación',
				links: [
					{ text: 'Primeros pasos', href: '/es/getting-started' },
					{ text: 'Arquitectura', href: '/es/architecture' },
					{ text: 'API del visor', href: '/es/api' },
					{ text: 'Registro de capacidades', href: '/parity' },
				],
			},
			{
				title: 'Comunidad',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio',
						external: true,
					},
					{
						text: 'Incidencias',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Licencia',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'visio-viewer · un visor de Visio temprano para la web',
	},
};
