import type { DefaultTheme } from 'vitepress';

const sidebar: DefaultTheme.SidebarItem[] = [
	{
		text: 'Pour commencer',
		items: [
			{ text: 'Vue d’ensemble', link: '/fr/' },
			{ text: 'Premiers pas', link: '/fr/getting-started' },
			{ text: 'Les démos en direct', link: '/demos' },
			{ text: 'Collaboration (démo en direct)', link: '/fr/collaboration' },
			{ text: 'Architecture', link: '/architecture' },
			{ text: 'Limites', link: '/fr/limitations' },
		],
	},
	{
		text: 'Référence',
		items: [
			{ text: 'API de l’élément et événements', link: '/fr/api' },
			{ text: 'Thèmes', link: '/theming' },
		],
	},
	{
		text: 'Guides par framework',
		items: [
			{ text: 'React', link: '/frameworks/react' },
			{ text: 'Vue', link: '/frameworks/vue' },
			{ text: 'Angular', link: '/frameworks/angular' },
			{ text: 'Svelte', link: '/frameworks/svelte' },
			{ text: 'Solid', link: '/frameworks/solid' },
			{ text: 'JS natif', link: '/frameworks/vanilla' },
		],
	},
	{
		text: 'Serveur',
		items: [
			{ text: 'Apportez votre propre serveur', link: '/fr/server' },
			{ text: 'Déployer le serveur de référence', link: '/fr/deploy' },
		],
	},
	{
		text: 'Démos',
		items: [
			{ text: 'Démo vanilla', link: '/demo/', target: '_self' },
			{ text: 'Démo React', link: '/demo-react/', target: '_self' },
			{ text: 'Démo Vue', link: '/demo-vue/', target: '_self' },
			{ text: 'Démo Angular', link: '/demo-angular/', target: '_self' },
			{ text: 'Démo Svelte', link: '/demo-svelte/', target: '_self' },
			{ text: 'Démo Solid', link: '/demo-solid/', target: '_self' },
			{ text: 'À propos des démos', link: '/demos' },
		],
	},
	{
		text: 'Projet',
		items: [{ text: 'Publication des paquets', link: '/releasing' }],
	},
];

export const fr = {
	label: 'Français',
	lang: 'fr-FR',
	link: '/fr/',
	description:
		'Un espace de travail d’équipe open source avec votre propre serveur : canaux et discussions, présence et appels WebRTC, avec des intégrations pour React, Vue, Angular, Svelte, Solid et JavaScript natif.',
	themeConfig: {
		langMenuLabel: 'Changer de langue',
		skipToContentLabel: 'Aller au contenu',
		darkModeSwitchLabel: 'Apparence',
		lightModeSwitchTitle: 'Passer au thème clair',
		darkModeSwitchTitle: 'Passer au thème sombre',
		sidebarMenuLabel: 'Menu',
		returnToTopLabel: 'Retour en haut',
		outline: { label: 'Sur cette page' },
		docFooter: { prev: 'Page précédente', next: 'Page suivante' },
		lastUpdated: { text: 'Dernière mise à jour' },
		editLink: {
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/teams/docs/:path',
			text: 'Modifier cette page sur GitHub',
		},
		footer: {
			message:
				'Publié sous licence Apache-2.0. OpenTeams n’est pas Microsoft Teams et n’est pas affilié à Microsoft.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		sidebar,
		nav: [
			{
				text: 'Guide',
				link: '/fr/getting-started',
				activeMatch: '/fr/(getting-started|limitations|api|collaboration|server|deploy)',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{ text: 'Versions', link: '/releasing', activeMatch: '/releasing' },
		],
		search: {
			provider: 'local',
			options: {
				translations: {
					button: { buttonText: 'Rechercher', buttonAriaLabel: 'Rechercher dans la documentation' },
					modal: {
						displayDetails: 'Afficher le détail',
						resetButtonTitle: 'Effacer la recherche',
						backButtonTitle: 'Fermer la recherche',
						noResultsText: 'Aucun résultat pour',
						footer: {
							selectText: 'sélectionner',
							selectKeyAriaLabel: 'Entrée',
							navigateText: 'naviguer',
							navigateUpKeyAriaLabel: 'Flèche haut',
							navigateDownKeyAriaLabel: 'Flèche bas',
							closeText: 'fermer',
							closeKeyAriaLabel: 'Échap',
						},
					},
				},
			},
		},
	} satisfies DefaultTheme.Config,
};
