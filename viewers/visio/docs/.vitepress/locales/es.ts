import type { DefaultTheme } from 'vitepress';

export const es = {
	label: 'Español',
	lang: 'es-ES',
	link: '/es/',
	description:
		'Un visor de diagramas de Visio para el navegador con un único componente web compartido, un único renderizador SVG y adaptadores ligeros para React, Vue, Angular, Svelte, Solid y JavaScript puro.',
	themeConfig: {
		langMenuLabel: 'Cambiar de idioma',
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
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/visio/docs/:path',
			text: 'Editar esta página en GitHub',
		},
		footer: {
			message: 'Publicado bajo la licencia Apache-2.0.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		nav: [
			{
				text: 'Guía',
				link: '/es/getting-started',
				activeMatch:
					'/es/(getting-started|architecture|api|collaboration)|/(getting-started|architecture|bindings|api|theming|collaboration|demos|frameworks)',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{ text: 'Versiones', link: '/releasing', activeMatch: '/releasing' },
		],
		sidebar: [
			{
				text: 'Comience aquí',
				items: [
					{ text: 'Resumen', link: '/es/' },
					{ text: 'Primeros pasos', link: '/es/getting-started' },
					{ text: 'Las demostraciones en vivo', link: '/demos' },
					{ text: 'Arquitectura', link: '/es/architecture' },
					{ text: 'Enlaces de frameworks', link: '/bindings' },
				],
			},
			{
				text: 'Referencia',
				items: [
					{ text: 'API del visor', link: '/es/api' },
					{ text: 'Temas', link: '/theming' },
				],
			},
			{
				text: 'Uso del visor',
				items: [{ text: 'Colaboración (uso compartido)', link: '/es/collaboration' }],
			},
			{
				text: 'Guías de frameworks',
				items: [
					{ text: 'React', link: '/frameworks/react' },
					{ text: 'Vue', link: '/frameworks/vue' },
					{ text: 'Angular', link: '/frameworks/angular' },
					{ text: 'Vanilla JS', link: '/frameworks/vanilla' },
					{ text: 'Svelte', link: '/frameworks/svelte' },
					{ text: 'Solid', link: '/frameworks/solid' },
				],
			},
			{
				text: 'Estado del proyecto',
				items: [
					{ text: 'Registro de capacidades', link: '/parity' },
					{ text: 'Registro de verificación', link: '/verification' },
					{ text: 'Comprobaciones del corpus', link: '/corpus' },
					{ text: 'Fuentes de la apariencia de la suite', link: '/suite-provenance' },
					{ text: 'Versiones de los paquetes', link: '/releasing' },
				],
			},
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
							navigateUpKeyAriaLabel: 'Flecha arriba',
							navigateDownKeyAriaLabel: 'Flecha abajo',
							closeText: 'cerrar',
							closeKeyAriaLabel: 'Esc',
						},
					},
				},
			},
		},
	} satisfies DefaultTheme.Config,
};
