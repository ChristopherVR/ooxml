import type { DefaultTheme } from 'vitepress';

const sidebar: DefaultTheme.SidebarItem[] = [
	{
		text: 'Pour commencer',
		items: [
			{ text: 'Vue d’ensemble', link: '/fr/' },
			{ text: 'Prise en main', link: '/fr/getting-started' },
			{ text: 'Les démos en direct', link: '/fr/demos' },
			{ text: 'Architecture', link: '/fr/architecture' },
			{ text: 'Liaisons de frameworks', link: '/fr/bindings' },
		],
	},
	{
		text: 'Référence',
		items: [
			{ text: 'API de l’élément', link: '/fr/api' },
			{ text: 'Thèmes', link: '/fr/theming' },
			{ text: 'Localisation', link: '/fr/localization' },
		],
	},
	{
		text: 'Guides par framework',
		items: [
			{ text: 'React', link: '/fr/frameworks/react' },
			{ text: 'Vue', link: '/fr/frameworks/vue' },
			{ text: 'Angular', link: '/fr/frameworks/angular' },
			{ text: 'JS vanilla', link: '/fr/frameworks/vanilla' },
			{ text: 'Svelte', link: '/fr/frameworks/svelte' },
			{ text: 'Solid', link: '/fr/frameworks/solid' },
		],
	},
	{
		text: 'État du projet',
		items: [
			{ text: 'Fonctionnalités et limites', link: '/fr/features' },
			{ text: 'Collaboration', link: '/fr/collaboration' },
			{ text: 'Publication des paquets', link: '/fr/releasing' },
			{
				text: 'Journal des modifications',
				link: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/CHANGELOG.md',
			},
		],
	},
];

export const fr = {
	label: 'Français',
	lang: 'fr-FR',
	link: '/fr/',
	description:
		'Un éditeur de tableur Excel dans le navigateur, avec un seul modèle de classeur, un seul composant web et des adaptateurs pour React, Vue, Angular, Svelte, Solid et JavaScript vanilla.',
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
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/xlsx/docs/:path',
			text: 'Modifier cette page sur GitHub',
		},
		footer: {
			message: 'Publié sous licence Apache-2.0.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		sidebar,
		nav: [
			{
				text: 'Guide',
				link: '/fr/getting-started',
				activeMatch:
					'/(fr/(getting-started|architecture|api|features|collaboration)|bindings|theming|localization|demos|frameworks)',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{ text: 'Versions', link: '/fr/releasing', activeMatch: '/releasing' },
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
