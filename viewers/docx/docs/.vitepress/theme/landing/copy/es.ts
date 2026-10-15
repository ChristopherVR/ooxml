import type { LandingCopy } from './types';

export const es: LandingCopy = {
	hero: {
		kicker: 'Implementación temprana · Apache-2.0 · TypeScript',
		titleTop: 'Edición de .docx,',
		titleAccent: 'lista para integrar.',
		sub: 'Un editor de Word para el navegador, para React, Vue, Angular, Svelte, Solid y JavaScript vanilla. Un único modelo de documento y un componente web <docx-editor> basado en ProseMirror hacen el trabajo; los adaptadores de framework solo conectan el ciclo de vida y los eventos.',
		start: { text: 'Empezar', href: '/es/getting-started' },
		demo: 'Demo en vivo',
		scroll: 'Desplazar',
		frameCaption: 'Documento de ejemplo.docx · en vivo en el navegador',
		frameTry: 'Pruébelo',
		frameAlt:
			'El editor docx-viewer mostrando un documento de ejemplo con la cinta de opciones al estilo de Word',
		copyLabel: 'Copiar',
		copiedLabel: 'Copiado',
		entryLabel: 'instalar',
		notPublished:
			'Publicado en npm. Cada paquete de framework es autónomo: incluye el editor y trae consigo docx-core.',
	},
	features: {
		kicker: 'Funciones',
		title: 'Lo que existe hoy.',
		items: [
			{
				title: 'Un único editor compartido',
				copy: 'Una cinta de opciones al estilo de Word, una barra de título, la vista Archivo y una barra de estado viven en un solo componente web, con deshacer/rehacer, buscar y reemplazar, zoom y un modo de solo lectura. Los adaptadores no añaden interfaz propia.',
				link: { text: 'Arquitectura', href: '/es/architecture' },
			},
			{
				title: 'Diseño de impresión',
				copy: 'Una vista paginada con salto de línea, reglas de viudas y huérfanas y de mantener juntos, columnas, división de filas de tabla, encabezados repetidos, colocación de notas al pie e impresión. Es una aproximación propia de este editor, no la paginación de Word.',
				link: { text: 'Hoja de ruta de compatibilidad', href: '/parity-roadmap' },
			},
			{
				title: 'Estructura del documento',
				copy: 'Estilos, listas numeradas y con viñetas, tablas, secciones y columnas, encabezados y pies de página, notas al pie y campos, incluida la tabla de contenido. Se pueden insertar cuadros de texto en línea, una marca de agua en el encabezado y bordes de página.',
				link: { text: 'Trabajo pendiente', href: '/outstanding-work' },
			},
			{
				title: 'Cambios con seguimiento y comentarios',
				copy: 'Modo Control de cambios, aceptar y rechazar, revisión siguiente y anterior, y un panel de comentarios con respuestas y resolución. Los comentarios se escriben de vuelta en comments.xml.',
				link: { text: 'Edición de texto', href: '/es/editing' },
			},
			{
				title: 'Preservación cuidadosa',
				copy: 'Un guardado sin cambios devuelve los bytes originales. Las ediciones compatibles conservan las partes del paquete ajenas al modelo, y las ediciones que dañarían contenido no compatible se rechazan en lugar de descartarse.',
				link: { text: 'Modelo y guardado', href: '/es/architecture' },
			},
			{
				title: '.doc heredado',
				copy: 'Los archivos de Word 97-2003 se cargan mediante los códecs compartidos de ole2. Se importa el texto del cuerpo principal, y las ediciones limitadas de párrafos existentes se guardan de nuevo como .doc. Los cambios de formato y de estructura se rechazan.',
				link: { text: 'Hoja de ruta de compatibilidad', href: '/parity-roadmap' },
			},
			{
				title: 'Protocolo de colaboración',
				copy: 'Un protocolo independiente del transporte intercambia pasos de ProseMirror validados a través de una autoridad, con presencia transitoria. Su aplicación se encarga de la red, la identidad, los permisos y el almacenamiento.',
				link: { text: 'Colaboración', href: '/es/collaboration' },
			},
			{
				title: 'Localización',
				copy: 'La interfaz del editor se entrega en inglés, francés, alemán, español y chino simplificado mediante una única opción de idioma. Traduce solo el texto de la interfaz, nunca el contenido del documento.',
				link: { text: 'Adaptadores', href: '/bindings' },
			},
		],
	},
	status: {
		kicker: 'Estado',
		title: 'Sinceros sobre lo que falta.',
		copy: 'Esta es una implementación temprana. No equivale a Microsoft Word: la paginación es aproximada, muchas funciones no se representan con la fidelidad de Word, los archivos protegidos con contraseña no son compatibles y Exportar DOCX no puede recuperar el contenido que el modelo omite. Las páginas de la hoja de ruta de compatibilidad y del trabajo pendiente detallan las carencias.',
		link: { text: 'Hoja de ruta de compatibilidad', href: '/parity-roadmap' },
		roadmapTitle: 'Nota de la hoja de ruta',
		roadmapCopy:
			'La lógica OOXML compartida se está trasladando al paquete privado ooxml-core, de modo que este repositorio pueda acabar conteniendo solo la interfaz. Esa migración está en curso y no es una función ya entregada.',
		roadmapLink: { text: 'El plan de ooxml-core', href: '/ooxml-core-plan' },
	},
	quickstart: {
		kicker: 'Primeros pasos',
		title: 'Monte un editor en unas pocas líneas.',
		copy: 'Todos los adaptadores montan el mismo <docx-editor>. Pase un modelo de documento, escuche los cambios y asigne una altura al contenedor.',
		docsLabel: 'Guía del framework',
		buildTitle: 'Compilar desde el código fuente',
		buildCopy:
			'Para ejecutar las demos o trabajar en el propio editor, clone el repositorio. Se necesita Bun.',
		buildCommands: 'bun install\nbun run demo',
	},
	demos: {
		kicker: 'Demo en vivo',
		title: 'Pruébelo aquí mismo.',
		copy: 'Este es el editor real ejecutándose en su navegador: la aplicación de demo compilada desde este repositorio para cada adaptador de framework, incrustada en vivo. Abra el documento de ejemplo o divida la vista para ver cómo dos pares locales comparten un documento.',
		frameworkLabel: 'Framework',
		languageLabel: 'Idioma',
		soloTab: 'Editor',
		collabTab: 'Colaboración',
		guestPicker: 'Framework de la ventana B',
		load: 'Cargar la demo en vivo',
		loading: 'Cargando el editor en vivo',
		openFull: 'Abrir la aplicación completa',
		hostLabel: 'Anfitrión (abre el ejemplo)',
		guestLabel: 'Se une por nombre de sesión',
		soloHint:
			'Use «Open the sample document» en la demo, o suelte un archivo DOCX o DOC propio. Cambiar de framework inicia un editor nuevo.',
		collabHint:
			'Dos ventanas, cada una en el framework que elija, se unen a una sesión por nombre. La ventana anfitriona ejecuta la autoridad de referencia en memoria; la otra recibe el documento y cada edición, junto con los cursores de ambos. La sesión es un BroadcastChannel, así que solo funciona entre ventanas de este navegador; no hay un servidor de red, y el transporte de producción corresponde a la aplicación anfitriona.',
	},
	faq: {
		kicker: 'Preguntas frecuentes',
		title: 'Dudas habituales.',
		items: [
			{
				q: '¿Puedo instalarlo desde npm?',
				a: 'Sí. docx-core y un paquete de editor autónomo por framework están en npm, por ejemplo npm install docx-react-viewer. Cada paquete tiene su propio versionado.',
				link: { text: 'Política de versiones', href: '/releasing' },
			},
			{
				q: '¿Es gratuito para uso comercial?',
				a: 'El repositorio se publica bajo la licencia Apache-2.0.',
			},
			{
				q: '¿Equivale a Word?',
				a: 'No. La paginación es una aproximación, muchas funciones no se representan con la fidelidad de Word y Exportar DOCX crea un documento nuevo a partir del contenido visible compatible. La equivalencia se sigue por separado para importación, diseño, edición, exportación, preservación y accesibilidad.',
				link: { text: 'Hoja de ruta de compatibilidad', href: '/parity-roadmap' },
			},
			{
				q: '¿Los archivos editados se vuelven a abrir en Word?',
				a: 'Las ediciones compatibles se escriben de vuelta en el paquete original, y las ediciones inseguras se rechazan en lugar de degradarlo en silencio. Es un objetivo de diseño comprobado con pruebas de regresión, no una afirmación de exportación sin pérdidas.',
				link: { text: 'Arquitectura', href: '/es/architecture' },
			},
			{
				q: '¿Qué frameworks son compatibles?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS y JavaScript sin framework. Comparten un único editor y modelo; cada adaptador solo se ocupa del montaje, la actualización de propiedades y el reenvío de eventos.',
				link: { text: 'Adaptadores', href: '/bindings' },
			},
			{
				q: '¿La colaboración necesita un servidor?',
				a: 'El protocolo es independiente del transporte e incluye una autoridad de referencia en memoria. La red, la autenticación, la persistencia y los permisos son responsabilidad de su aplicación.',
				link: { text: 'Colaboración', href: '/es/collaboration' },
			},
			{
				q: '¿Puede abrir archivos .doc antiguos?',
				a: 'Sí, con límites. Se importa el texto del cuerpo principal y las ediciones limitadas de párrafos existentes se guardan de nuevo como .doc; los cambios de formato y de estructura se rechazan.',
			},
			{
				q: '¿Y los archivos protegidos con contraseña?',
				a: 'Hoy no son compatibles.',
			},
		],
	},
	finale: {
		kicker: 'Empezar',
		title: '.docx de entrada. .docx de salida.',
		sub: 'Lea la arquitectura, elija la guía de su framework y pruebe la demo con un documento propio. Licencia Apache-2.0, TypeScript estricto y sinceridad sobre las carencias.',
		quick: { text: 'Leer la guía', href: '/es/getting-started' },
		github: 'Ver en GitHub',
		columns: [
			{
				title: 'Producto',
				links: [
					{
						text: 'Demo en vivo',
						href: 'https://christophervr.github.io/ooxml/docx/demo/',
						external: true,
					},
					{ text: 'Adaptadores de framework', href: '/bindings' },
					{ text: 'Colaboración', href: '/es/collaboration' },
					{ text: 'Versiones', href: '/releasing' },
				],
			},
			{
				title: 'Documentación',
				links: [
					{ text: 'Primeros pasos', href: '/es/getting-started' },
					{ text: 'Arquitectura', href: '/es/architecture' },
					{ text: 'Edición de texto', href: '/es/editing' },
					{ text: 'Hoja de ruta de compatibilidad', href: '/parity-roadmap' },
					{ text: 'Trabajo pendiente', href: '/outstanding-work' },
				],
			},
			{
				title: 'Comunidad',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx',
						external: true,
					},
					{
						text: 'Incidencias',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Licencia',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'docx-viewer · un editor de Word temprano para la web',
	},
};
