import type { DefaultTheme } from 'vitepress';

export const es = {
	label: 'Español',
	lang: 'es-ES',
	link: '/es/',
	description:
		'Una base de edición de documentos Word en el navegador, con un único modelo de documento, un editor como componente web y adaptadores para React, Vue, Angular, Svelte, Solid y JavaScript vanilla.',
	themeConfig: {
		langMenuLabel: 'Cambiar idioma',
		skipToContentLabel: 'Saltar al contenido',
		darkModeSwitchLabel: 'Apariencia',
		lightModeSwitchTitle: 'Cambiar al tema claro',
		darkModeSwitchTitle: 'Cambiar al tema oscuro',
		sidebarMenuLabel: 'Menú',
		returnToTopLabel: 'Volver arriba',
		outline: { label: 'En esta página' },
		docFooter: { prev: 'Página anterior', next: 'Página siguiente' },
		lastUpdated: { text: 'Última actualización' },
		editLink: {
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/docx/docs/:path',
			text: 'Editar esta página en GitHub',
		},
		footer: {
			message: 'Publicado bajo la licencia Apache-2.0.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		sidebar: [
			{
				text: 'Primeros pasos',
				items: [
					{ text: 'Resumen', link: '/es/' },
					{ text: 'Primeros pasos', link: '/es/getting-started' },
					{ text: 'Las demos en vivo', link: '/es/demos' },
					{ text: 'Arquitectura', link: '/es/architecture' },
					{ text: 'Adaptadores de framework', link: '/es/bindings' },
					{ text: 'Bloques de construcción', link: '/es/building-blocks' },
					{ text: 'Unidades del modelo', link: '/es/model-units' },
				],
			},
			{
				text: 'Referencia',
				items: [
					{ text: 'API del elemento y eventos', link: '/es/api' },
					{ text: 'Temas', link: '/es/theming' },
				],
			},
			{
				text: 'Uso del editor',
				items: [
					{ text: 'Edición de texto', link: '/es/editing' },
					{ text: 'Colaboración (demo en vivo)', link: '/es/collaboration' },
				],
			},
			{
				text: 'Guías por framework',
				items: [
					{ text: 'React', link: '/es/frameworks/react' },
					{ text: 'Vue', link: '/es/frameworks/vue' },
					{ text: 'Angular', link: '/es/frameworks/angular' },
					{ text: 'Vanilla JS', link: '/es/frameworks/vanilla' },
					{ text: 'Svelte', link: '/es/frameworks/svelte' },
					{ text: 'Solid', link: '/es/frameworks/solid' },
				],
			},
			{
				text: 'Estado del proyecto',
				items: [
					{ text: 'Hoja de ruta de compatibilidad', link: '/parity-roadmap' },
					{ text: 'Trabajo pendiente', link: '/outstanding-work' },
					{ text: 'Publicación de paquetes', link: '/es/releasing' },
				],
			},
		],
		nav: [
			{
				text: 'Guía',
				link: '/es/getting-started',
				activeMatch:
					'/(es/)?(getting-started|architecture|bindings|model-units|api|theming|editing|collaboration|demos|frameworks)',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{ text: 'Versiones', link: '/es/releasing', activeMatch: '/releasing' },
		],
		search: {
			provider: 'local',
			options: {
				translations: {
					button: { buttonText: 'Buscar', buttonAriaLabel: 'Buscar en la documentación' },
					modal: {
						displayDetails: 'Mostrar detalles',
						resetButtonTitle: 'Borrar la búsqueda',
						backButtonTitle: 'Cerrar la búsqueda',
						noResultsText: 'No se encontraron resultados para',
						footer: {
							selectText: 'seleccionar',
							selectKeyAriaLabel: 'Intro',
							navigateText: 'navegar',
							navigateUpKeyAriaLabel: 'flecha arriba',
							navigateDownKeyAriaLabel: 'flecha abajo',
							closeText: 'cerrar',
							closeKeyAriaLabel: 'Esc',
						},
					},
				},
			},
		},
	} satisfies DefaultTheme.Config,
};
