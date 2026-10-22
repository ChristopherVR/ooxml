import type { DefaultTheme } from 'vitepress';

export const fr = {
	label: 'Français',
	lang: 'fr-FR',
	link: '/fr/',
	description:
		'Un visualiseur de diagrammes Visio pour le navigateur, avec un composant web partagé, un moteur de rendu SVG unique et de fins adaptateurs pour React, Vue, Angular, Svelte, Solid et JavaScript vanilla.',
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
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/visio/docs/:path',
			text: 'Modifier cette page sur GitHub',
		},
		footer: {
			message: 'Publié sous licence Apache-2.0.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		nav: [
			{
				text: 'Guide',
				link: '/fr/getting-started',
				activeMatch:
					'/fr/(getting-started|architecture|api|collaboration)|/(getting-started|architecture|bindings|api|theming|building-blocks|collaboration|demos|frameworks)',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{ text: 'Versions', link: '/releasing', activeMatch: '/releasing' },
		],
		sidebar: [
			{
				text: 'Pour commencer',
				items: [
					{ text: 'Vue d’ensemble', link: '/fr/' },
					{ text: 'Prise en main', link: '/fr/getting-started' },
					{ text: 'Les démos en direct', link: '/demos' },
					{ text: 'Architecture', link: '/fr/architecture' },
					{ text: 'Liaisons de frameworks', link: '/bindings' },
				],
			},
			{
				text: 'Référence',
				items: [
					{ text: 'API du visualiseur', link: '/fr/api' },
					{ text: 'Thèmes', link: '/theming' },
				],
			},
			{
				text: 'Utiliser le visualiseur',
				items: [
					{ text: 'Blocs de construction', link: '/building-blocks' },
					{ text: 'Collaboration (partage)', link: '/fr/collaboration' },
				],
			},
			{
				text: 'Guides par framework',
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
				text: 'État du projet',
				items: [
					{ text: 'Registre des capacités', link: '/parity' },
					{ text: 'Registre de vérification', link: '/verification' },
					{ text: 'Vérifications du corpus', link: '/corpus' },
					{ text: 'Sources de l’apparence de la suite', link: '/suite-provenance' },
					{ text: 'Versions des paquets', link: '/releasing' },
				],
			},
		],
		search: {
			provider: 'local',
			options: {
				translations: {
					button: { buttonText: 'Rechercher', buttonAriaLabel: 'Rechercher dans la documentation' },
					modal: {
						displayDetails: 'Afficher les détails',
						resetButtonTitle: 'Effacer la recherche',
						backButtonTitle: 'Fermer la recherche',
						noResultsText: 'Aucun résultat pour',
						footer: {
							selectText: 'sélectionner',
							selectKeyAriaLabel: 'entrée',
							navigateText: 'naviguer',
							navigateUpKeyAriaLabel: 'flèche vers le haut',
							navigateDownKeyAriaLabel: 'flèche vers le bas',
							closeText: 'fermer',
							closeKeyAriaLabel: 'échap',
						},
					},
				},
			},
		},
	} satisfies DefaultTheme.Config,
};
