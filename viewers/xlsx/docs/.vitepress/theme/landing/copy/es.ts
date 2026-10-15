import type { LandingCopy } from './types';

export const es: LandingCopy = {
	hero: {
		kicker: 'Implementación temprana · Apache-2.0 · TypeScript',
		titleTop: 'Edición de .xlsx,',
		titleAccent: 'lista para incrustar.',
		sub: 'Un editor de hojas de cálculo para el navegador con React, Vue, Angular, Svelte, Solid y JavaScript vanilla. Un único modelo de libro con motor de fórmulas y un único componente web <xlsx-editor> hacen el trabajo; los adaptadores de framework solo conectan el ciclo de vida y los eventos.',
		start: { text: 'Primeros pasos', href: '/es/getting-started' },
		demo: 'Demo en vivo',
		scroll: 'Desplazar',
		frameCaption: 'Libro de ejemplo.xlsx · en vivo en el navegador',
		frameTry: 'Pruébelo',
		frameAlt:
			'Ilustración de la cuadrícula de xlsx-viewer con cinta de opciones, barra de fórmulas y pestañas de hoja',
		copyLabel: 'Copiar',
		copiedLabel: 'Copiado',
		entryLabel: 'instalar',
		notPublished:
			'Publicado en npm. Cada paquete de framework es autónomo: incluye el editor y trae consigo el modelo de libro.',
	},
	features: {
		kicker: 'Funciones',
		title: 'Para qué está construido.',
		items: [
			{
				title: 'Un único editor compartido',
				copy: 'Una cinta de opciones, un cuadro de nombres, una barra de fórmulas, la cuadrícula, las pestañas de hoja y una barra de estado viven en un solo componente web, con deshacer y rehacer y un modo de solo lectura. Los enlaces no añaden interfaz propia.',
				link: { text: 'API del elemento', href: '/es/api' },
			},
			{
				title: 'Fórmulas que se recalculan',
				copy: 'Un motor de fórmulas en ooxml-core analiza y evalúa las fórmulas de las celdas y recalcula las dependientes tras cada edición. Las funciones desconocidas muestran #NAME? en lugar de un resultado inventado.',
				link: { text: 'Funciones y limitaciones', href: '/es/features' },
			},
			{
				title: 'Formato tal como se guardó',
				copy: 'Las fuentes, los rellenos, los bordes, la alineación, los formatos numéricos, las celdas combinadas, los paneles inmovilizados, las anchuras de columna y las alturas de fila se leen del archivo y los dibuja la cuadrícula.',
				link: { text: 'Funciones y limitaciones', href: '/es/features' },
			},
			{
				title: 'Formatos condicionales y validación',
				copy: 'Las escalas de color, las barras de datos, los conjuntos de iconos y las reglas de celda se evalúan sobre valores en vivo; las validaciones de lista desplegable ofrecen sus opciones en la cuadrícula.',
				link: { text: 'Funciones y limitaciones', href: '/es/features' },
			},
			{
				title: 'Gráficos, para visualizar',
				copy: 'Los tipos de gráfico comunes se dibujan como SVG a partir de valores de celda en vivo. No se ofrece la edición de gráficos; las partes del gráfico se escriben de nuevo sin cambios.',
				link: { text: 'Funciones y limitaciones', href: '/es/features' },
			},
			{
				title: '.xls heredado y CSV',
				copy: 'Los libros de Excel 97-2003 se cargan mediante los códecs compartidos de ole2 incluidos en ooxml-core y se guardan como .xlsx. El CSV se abre y se exporta desde la hoja activa.',
				link: { text: 'Funciones y limitaciones', href: '/es/features' },
			},
			{
				title: 'Conservación cuidadosa',
				copy: 'Las partes que el modelo no cubre (proyectos VBA, cachés dinámicas, XML personalizado, detalles de gráficos) se transmiten al guardar mientras la hoja siga existiendo.',
				link: { text: 'Arquitectura', href: '/es/architecture' },
			},
			{
				title: 'Localización',
				copy: 'La interfaz está preparada para inglés, francés, alemán, español y chino simplificado mediante una sola opción de configuración regional. El contenido del libro nunca se traduce.',
				link: { text: 'Localización', href: '/localization' },
			},
		],
	},
	status: {
		kicker: 'Estado',
		title: 'Sincero sobre lo que falta.',
		copy: 'Esta es una implementación temprana y no tiene paridad con Microsoft Excel. Las tablas dinámicas, las segmentaciones de datos, las macros y los vínculos externos se conservan pero no se muestran ni se editan; los minigráficos se muestran pero no se editan; la biblioteca de funciones es amplia pero no está completa; los archivos protegidos con contraseña no están admitidos. La página de funciones enumera las carencias.',
		link: { text: 'Funciones y limitaciones', href: '/es/features' },
		roadmapTitle: 'Dónde vive la lógica',
		roadmapCopy:
			'El modelo de libro, el motor de fórmulas, el analizador, el serializador y el lector de .xls viven en el área xlsx del paquete público ooxml-core. Este repositorio contiene únicamente la interfaz.',
		roadmapLink: { text: 'Arquitectura', href: '/es/architecture' },
	},
	quickstart: {
		kicker: 'Primeros pasos',
		title: 'Monte un editor en pocas líneas.',
		copy: 'Cada adaptador monta el mismo <xlsx-editor>. Pase un libro o los bytes de un archivo, escuche los cambios y dé una altura al contenedor.',
		docsLabel: 'Guía del framework',
		buildTitle: 'Compilar desde el código fuente',
		buildCopy:
			'Para ejecutar las demos o trabajar en el propio editor, clone el repositorio. Se necesita Bun.',
		buildCommands: 'bun install\nbun run demo',
	},
	demos: {
		kicker: 'Demo en vivo',
		title: 'Pruébelo aquí mismo.',
		copy: 'Este es el editor ejecutándose en su navegador: la aplicación de demo compilada desde este repositorio para cada adaptador de framework, incrustada en vivo con el libro de ejemplo abierto.',
		frameworkLabel: 'Framework',
		languageLabel: 'Idioma',
		load: 'Cargar la demo en vivo',
		loading: 'Cargando el editor en vivo',
		openFull: 'Abrir la aplicación completa',
		hint: 'El ejemplo tiene tres hojas con fórmulas, formato condicional, una lista de validación y gráficos. Abra la aplicación completa para soltar un archivo .xlsx, .xls o .csv propio. Cambiar de framework inicia un editor nuevo.',
	},
	faq: {
		kicker: 'Preguntas frecuentes',
		title: 'Preguntas habituales.',
		items: [
			{
				q: '¿Puedo instalarlo desde npm?',
				a: 'Sí. Hay un paquete de editor autónomo por framework en npm, por ejemplo npm install @christophervr/xlsx-react-viewer. Cada paquete tiene su propio versionado.',
				link: { text: 'Política de publicación', href: '/releasing' },
			},
			{
				q: '¿Se puede usar gratis con fines comerciales?',
				a: 'El repositorio tiene licencia Apache-2.0.',
			},
			{
				q: '¿Tiene paridad con Excel?',
				a: 'No. La cuadrícula dibuja lo que el modelo entiende, el motor de fórmulas cubre un conjunto amplio pero incompleto de funciones y varias características de los libros se conservan sin mostrarse ni editarse.',
				link: { text: 'Funciones y limitaciones', href: '/es/features' },
			},
			{
				q: '¿Los archivos guardados se abren en Excel?',
				a: 'Se escriben para abrirse sin reparación, y las pruebas de ooxml-core comprueban la ida y vuelta. Es un objetivo de diseño, no una afirmación de exportación sin pérdidas.',
				link: { text: 'Arquitectura', href: '/es/architecture' },
			},
			{
				q: '¿Qué frameworks están admitidos?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS y JavaScript sin framework. Comparten un único editor y modelo; cada adaptador solo se ocupa del montaje, la actualización de propiedades y el reenvío de eventos.',
				link: { text: 'Enlaces de framework', href: '/bindings' },
			},
			{
				q: '¿Pueden varias personas editar un mismo libro?',
				a: 'Todavía no. La colaboración en tiempo real no está implementada para hojas de cálculo.',
				link: { text: 'Colaboración', href: '/es/collaboration' },
			},
			{
				q: '¿Puede abrir archivos .xls antiguos?',
				a: 'Sí, en solo lectura: se cargan los valores, las fórmulas que se pueden decodificar, el formato básico, las celdas combinadas y el diseño de las hojas, y el libro se guarda como .xlsx.',
			},
			{
				q: '¿Y los archivos protegidos con contraseña?',
				a: 'No están admitidos por ahora.',
			},
		],
	},
	finale: {
		kicker: 'Empiece ya',
		title: '.xlsx dentro. .xlsx fuera.',
		sub: 'Lea la guía de primeros pasos, elija su framework y pruebe la demo con un libro propio. Licencia Apache-2.0, TypeScript estricto y sinceridad sobre las carencias.',
		quick: { text: 'Leer la guía', href: '/es/getting-started' },
		github: 'Ver en GitHub',
		columns: [
			{
				title: 'Producto',
				links: [
					{
						text: 'Demo en vivo',
						href: 'https://christophervr.github.io/ooxml/xlsx/demo/',
						external: true,
					},
					{ text: 'Enlaces de framework', href: '/bindings' },
					{ text: 'API del elemento', href: '/es/api' },
					{ text: 'Versiones', href: '/releasing' },
				],
			},
			{
				title: 'Documentación',
				links: [
					{ text: 'Primeros pasos', href: '/es/getting-started' },
					{ text: 'Demos en vivo', href: '/demos' },
					{ text: 'Arquitectura', href: '/es/architecture' },
					{ text: 'Temas', href: '/theming' },
					{ text: 'Funciones y limitaciones', href: '/es/features' },
				],
			},
			{
				title: 'Comunidad',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx',
						external: true,
					},
					{
						text: 'Incidencias',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Licencia',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'xlsx-viewer · un editor de hojas de cálculo temprano para la web',
	},
};
