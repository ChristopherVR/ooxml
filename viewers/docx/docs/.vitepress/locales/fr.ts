import type { DefaultTheme } from 'vitepress';

export const fr = {
	label: 'Français',
	lang: 'fr-FR',
	link: '/fr/',
	description:
		"Une base d'édition de documents Word dans le navigateur, avec un seul modèle de document, un éditeur en composant web et des adaptateurs pour React, Vue, Angular, Svelte, Solid et JavaScript vanilla.",
	themeConfig: {
		langMenuLabel: 'Changer de langue',
		skipToContentLabel: 'Passer au contenu',
		darkModeSwitchLabel: 'Apparence',
		lightModeSwitchTitle: 'Passer au thème clair',
		darkModeSwitchTitle: 'Passer au thème sombre',
		sidebarMenuLabel: 'Menu',
		returnToTopLabel: 'Retour en haut',
		outline: { label: 'Sur cette page' },
		docFooter: { prev: 'Page précédente', next: 'Page suivante' },
		lastUpdated: { text: 'Dernière mise à jour' },
		editLink: {
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/docx/docs/:path',
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
					'/(fr/)?(getting-started|architecture|bindings|model-units|api|theming|editing|collaboration|demos|frameworks)',
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
					{ text: 'Adaptateurs de frameworks', link: '/bindings' },
					{ text: 'Unités du modèle', link: '/model-units' },
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
				text: 'Utiliser l’éditeur',
				items: [
					{ text: 'Édition du texte', link: '/fr/editing' },
					{ text: 'Collaboration (démo en direct)', link: '/fr/collaboration' },
				],
			},
			{
				text: 'Guides par framework',
				items: [
					{ text: 'React', link: '/frameworks/react' },
					{ text: 'Vue', link: '/frameworks/vue' },
					{ text: 'Angular', link: '/frameworks/angular' },
					{ text: 'JS vanilla', link: '/frameworks/vanilla' },
					{ text: 'Svelte', link: '/frameworks/svelte' },
					{ text: 'Solid', link: '/frameworks/solid' },
				],
			},
			{
				text: 'État du projet',
				items: [
					{ text: 'Feuille de route de la prise en charge', link: '/parity-roadmap' },
					{ text: 'Travaux restants', link: '/outstanding-work' },
					{ text: 'Plan d’ooxml-core', link: '/ooxml-core-plan' },
					{ text: 'Code de format partagé', link: '/reuse-audit' },
					{ text: 'Publication des paquets', link: '/releasing' },
				],
			},
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
							selectKeyAriaLabel: 'entrée',
							navigateText: 'naviguer',
							navigateUpKeyAriaLabel: 'flèche du haut',
							navigateDownKeyAriaLabel: 'flèche du bas',
							closeText: 'fermer',
							closeKeyAriaLabel: 'échap',
						},
					},
				},
			},
		},
	} satisfies DefaultTheme.Config,
};
