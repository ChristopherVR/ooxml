import type { DefaultTheme } from 'vitepress';

const sidebar: DefaultTheme.SidebarItem[] = [
	{
		text: 'Empezar',
		items: [
			{ text: 'Resumen', link: '/es/' },
			{ text: 'Primeros pasos', link: '/es/getting-started' },
			{ text: 'Las demos en vivo', link: '/es/demos' },
			{ text: 'Arquitectura', link: '/es/architecture' },
			{ text: 'Enlaces de framework', link: '/es/bindings' },
		],
	},
	{
		text: 'Referencia',
		items: [
			{ text: 'API del elemento', link: '/es/api' },
			{ text: 'Temas', link: '/es/theming' },
			{ text: 'Localización', link: '/es/localization' },
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
			{ text: 'Funciones y limitaciones', link: '/es/features' },
			{ text: 'Colaboración', link: '/es/collaboration' },
			{ text: 'Publicación de paquetes', link: '/es/releasing' },
			{
				text: 'Registro de cambios',
				link: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/CHANGELOG.md',
			},
		],
	},
];

export const es = {
	label: 'Español',
	lang: 'es-ES',
	link: '/es/',
	description:
		'Un editor de hojas de cálculo de Excel para el navegador, con un único modelo de libro, un editor como componente web y adaptadores para React, Vue, Angular, Svelte, Solid y JavaScript vanilla.',
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
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/xlsx/docs/:path',
			text: 'Editar esta página en GitHub',
		},
		footer: {
			message: 'Publicado bajo la licencia Apache-2.0.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		sidebar,
		nav: [
			{
				text: 'Guía',
				link: '/es/getting-started',
				activeMatch:
					'/(es/(getting-started|architecture|api|features|collaboration)|getting-started|architecture|bindings|api|theming|localization|demos|features|collaboration|frameworks)',
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
						noResultsText: 'No se encontraron resultados',
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
