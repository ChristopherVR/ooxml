import type { DefaultTheme } from 'vitepress';

const sidebar: DefaultTheme.SidebarItem[] = [
	{
		text: 'Empieza aquí',
		items: [
			{ text: 'Resumen', link: '/es/' },
			{ text: 'Primeros pasos', link: '/es/getting-started' },
			{ text: 'Las demos en vivo', link: '/demos' },
			{ text: 'Colaboración (demo en vivo)', link: '/es/collaboration' },
			{ text: 'Arquitectura', link: '/architecture' },
			{ text: 'Limitaciones', link: '/es/limitations' },
		],
	},
	{
		text: 'Referencia',
		items: [
			{ text: 'API del elemento y eventos', link: '/es/api' },
			{ text: 'Temas', link: '/theming' },
		],
	},
	{
		text: 'Guías por framework',
		items: [
			{ text: 'React', link: '/frameworks/react' },
			{ text: 'Vue', link: '/frameworks/vue' },
			{ text: 'Angular', link: '/frameworks/angular' },
			{ text: 'Svelte', link: '/frameworks/svelte' },
			{ text: 'Solid', link: '/frameworks/solid' },
			{ text: 'JS puro', link: '/frameworks/vanilla' },
		],
	},
	{
		text: 'Servidor',
		items: [
			{ text: 'Aporta tu propio servidor', link: '/es/server' },
			{ text: 'Despliegue del servidor de referencia', link: '/es/deploy' },
		],
	},
	{
		text: 'Demos',
		items: [
			// `target` keeps the VitePress router from handling these as missing pages.
			{ text: 'Demo vanilla', link: '/demo/', target: '_self' },
			{ text: 'Demo de React', link: '/demo-react/', target: '_self' },
			{ text: 'Demo de Vue', link: '/demo-vue/', target: '_self' },
			{ text: 'Demo de Angular', link: '/demo-angular/', target: '_self' },
			{ text: 'Demo de Svelte', link: '/demo-svelte/', target: '_self' },
			{ text: 'Demo de Solid', link: '/demo-solid/', target: '_self' },
			{ text: 'Acerca de las demos', link: '/demos' },
		],
	},
	{
		text: 'Proyecto',
		items: [{ text: 'Publicación de paquetes', link: '/releasing' }],
	},
];

export const es = {
	label: 'Español',
	lang: 'es-ES',
	link: '/es/',
	description:
		'Un espacio de trabajo de equipo de código abierto en el que aportas tu propio servidor: canales y chat, presencia y llamadas WebRTC, con bindings para React, Vue, Angular, Svelte, Solid y JavaScript puro.',
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
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/teams/docs/:path',
			text: 'Editar esta página en GitHub',
		},
		footer: {
			message:
				'Publicado bajo la licencia Apache-2.0. OpenTeams no es Microsoft Teams y no está afiliado a Microsoft.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		sidebar,
		nav: [
			{
				text: 'Guía',
				link: '/es/getting-started',
				activeMatch:
					'/(es/(getting-started|limitations|api|collaboration|server|deploy)|getting-started|architecture|limitations|api|theming|collaboration|server|deploy)',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{ text: 'Versiones', link: '/releasing', activeMatch: '/releasing' },
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
							selectKeyAriaLabel: 'intro',
							navigateText: 'navegar',
							navigateUpKeyAriaLabel: 'flecha arriba',
							navigateDownKeyAriaLabel: 'flecha abajo',
							closeText: 'cerrar',
							closeKeyAriaLabel: 'esc',
						},
					},
				},
			},
		},
	} satisfies DefaultTheme.Config,
};
