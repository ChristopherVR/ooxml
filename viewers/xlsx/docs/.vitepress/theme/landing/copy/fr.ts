import type { LandingCopy } from './types';

export const fr: LandingCopy = {
	hero: {
		kicker: 'Implémentation précoce · Apache-2.0 · TypeScript',
		titleTop: 'Édition .xlsx,',
		titleAccent: 'prête à intégrer.',
		sub: 'Un éditeur de tableur dans le navigateur pour React, Vue, Angular, Svelte, Solid et JavaScript vanilla. Un seul modèle de classeur avec un moteur de formules et un seul composant web <xlsx-editor> font le travail ; les adaptateurs de frameworks ne gèrent que le cycle de vie et les événements.',
		start: { text: 'Commencer', href: '/fr/getting-started' },
		demo: 'Démo en direct',
		scroll: 'Défiler',
		frameCaption: 'Classeur d’exemple .xlsx · en direct dans le navigateur',
		frameTry: 'Essayer',
		frameAlt:
			'Une illustration de la grille xlsx-viewer avec un ruban, une barre de formule et des onglets de feuille',
		copyLabel: 'Copier',
		copiedLabel: 'Copié',
		entryLabel: 'installer',
		notPublished:
			'Publié sur npm. Chaque paquet de framework est autonome : il embarque l’éditeur et apporte avec lui le modèle de classeur.',
	},
	features: {
		kicker: 'Fonctionnalités',
		title: 'Ce pour quoi il est conçu.',
		items: [
			{
				title: 'Un seul éditeur partagé',
				copy: 'Un ruban, une zone Nom, une barre de formule, une grille, des onglets de feuille et une barre d’état vivent dans un seul composant web, avec annuler et rétablir et un mode lecture seule. Les liaisons n’ajoutent aucune interface.',
				link: { text: 'API de l’élément', href: '/fr/api' },
			},
			{
				title: 'Des formules qui se recalculent',
				copy: 'Un moteur de formules dans ooxml-core analyse et évalue les formules de cellule, en recalculant les cellules dépendantes après chaque modification. Les fonctions inconnues affichent #NAME? au lieu d’un résultat deviné.',
				link: { text: 'Fonctionnalités et limites', href: '/fr/features' },
			},
			{
				title: 'La mise en forme telle qu’enregistrée',
				copy: 'Polices, remplissages, bordures, alignement, formats numériques, cellules fusionnées, volets figés, largeurs de colonne et hauteurs de ligne sont lus dans le fichier et dessinés par la grille.',
				link: { text: 'Fonctionnalités et limites', href: '/fr/features' },
			},
			{
				title: 'Mises en forme conditionnelles et validation',
				copy: 'Les échelles de couleurs, barres de données, jeux d’icônes et règles de cellule sont évalués sur les valeurs en direct ; les validations par liste déroulante proposent leurs choix dans la grille.',
				link: { text: 'Fonctionnalités et limites', href: '/fr/features' },
			},
			{
				title: 'Des graphiques affichés',
				copy: 'Les types de graphiques courants sont dessinés en SVG à partir des valeurs de cellule en direct. La modification des graphiques n’est pas proposée ; les parties de graphique sont réécrites telles quelles.',
				link: { text: 'Fonctionnalités et limites', href: '/fr/features' },
			},
			{
				title: '.xls ancien format et CSV',
				copy: 'Les classeurs Excel 97-2003 se chargent via les codecs ole2 partagés dans ooxml-core et s’enregistrent en .xlsx. Le CSV s’ouvre et s’exporte depuis la feuille active.',
				link: { text: 'Fonctionnalités et limites', href: '/fr/features' },
			},
			{
				title: 'Une conservation soignée',
				copy: 'Les parties que le modèle ne couvre pas (projets VBA, caches de tableaux croisés dynamiques, XML personnalisé, détails des graphiques) sont reportées à l’enregistrement tant que la feuille existe.',
				link: { text: 'Architecture', href: '/fr/architecture' },
			},
			{
				title: 'Localisation',
				copy: 'L’interface est prévue pour l’anglais, le français, l’allemand, l’espagnol et le chinois simplifié via une seule option de langue. Le contenu des classeurs n’est jamais traduit.',
				link: { text: 'Localisation', href: '/localization' },
			},
		],
	},
	status: {
		kicker: 'État',
		title: 'Honnête sur ce qui n’est pas fait.',
		copy: 'Il s’agit d’une implémentation précoce, pas de la parité avec Microsoft Excel. Les tableaux croisés dynamiques, segments, macros et liaisons externes sont conservés mais ni affichés ni modifiés ; les sparklines sont affichés mais non modifiés ; la bibliothèque de fonctions est vaste mais incomplète ; les fichiers protégés par mot de passe ne sont pas pris en charge. La page des fonctionnalités liste les écarts.',
		link: { text: 'Fonctionnalités et limites', href: '/fr/features' },
		roadmapTitle: 'Où vit la logique',
		roadmapCopy:
			'Le modèle de classeur, le moteur de formules, l’analyseur, le sérialiseur et le lecteur .xls vivent dans la zone xlsx du paquet public ooxml-core. Ce dépôt ne contient que l’interface.',
		roadmapLink: { text: 'Architecture', href: '/fr/architecture' },
	},
	quickstart: {
		kicker: 'Prise en main',
		title: 'Montez un éditeur en quelques lignes.',
		copy: 'Chaque adaptateur monte le même <xlsx-editor>. Passez un classeur ou les octets d’un fichier, écoutez les modifications et donnez une hauteur au conteneur.',
		docsLabel: 'Guide du framework',
		buildTitle: 'Compiler depuis les sources',
		buildCopy:
			'Pour lancer les démos ou travailler sur l’éditeur lui-même, clonez le dépôt. Bun est requis.',
		buildCommands: 'bun install\nbun run demo',
	},
	demos: {
		kicker: 'Démo en direct',
		title: 'Essayez-le ici même.',
		copy: 'Voici l’éditeur qui tourne dans votre navigateur : l’application de démo compilée depuis ce dépôt pour chaque adaptateur de framework, intégrée en direct avec le classeur d’exemple ouvert.',
		frameworkLabel: 'Framework',
		languageLabel: 'Langue',
		load: 'Charger la démo en direct',
		loading: 'Chargement de l’éditeur en direct',
		openFull: 'Ouvrir l’application complète',
		hint: 'L’exemple contient trois feuilles avec des formules, des mises en forme conditionnelles, une liste de validation et des graphiques. Ouvrez l’application complète pour y déposer un fichier .xlsx, .xls ou .csv de votre choix. Changer de framework démarre un nouvel éditeur.',
	},
	faq: {
		kicker: 'FAQ',
		title: 'Questions fréquentes.',
		items: [
			{
				q: 'Puis-je l’installer depuis npm ?',
				a: 'Oui. Un paquet d’éditeur autonome par framework est sur npm, par exemple npm install @christophervr/xlsx-react-viewer. Chaque paquet est versionné séparément.',
				link: { text: 'Politique de publication', href: '/releasing' },
			},
			{
				q: 'Est-il gratuit pour un usage commercial ?',
				a: 'Le dépôt est sous licence Apache-2.0.',
			},
			{
				q: 'Est-ce la parité avec Excel ?',
				a: 'Non. La grille dessine ce que le modèle comprend, le moteur de formules couvre un ensemble de fonctions vaste mais incomplet, et plusieurs fonctionnalités de classeur sont conservées sans être affichées ni modifiées.',
				link: { text: 'Fonctionnalités et limites', href: '/fr/features' },
			},
			{
				q: 'Les fichiers enregistrés s’ouvrent-ils dans Excel ?',
				a: 'Ils sont écrits pour s’ouvrir sans réparation, et l’aller-retour est vérifié par des tests dans ooxml-core. C’est un objectif de conception, pas une promesse d’export sans perte.',
				link: { text: 'Architecture', href: '/fr/architecture' },
			},
			{
				q: 'Quels frameworks sont pris en charge ?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS et JavaScript pur. Ils partagent un seul éditeur et un seul modèle ; chaque adaptateur ne gère que le montage, la mise à jour des propriétés et la transmission des événements.',
				link: { text: 'Liaisons de frameworks', href: '/bindings' },
			},
			{
				q: 'Plusieurs personnes peuvent-elles modifier un même classeur ?',
				a: 'Pas encore. La collaboration en temps réel n’est pas implémentée pour les tableurs.',
				link: { text: 'Collaboration', href: '/fr/collaboration' },
			},
			{
				q: 'Peut-il ouvrir d’anciens fichiers .xls ?',
				a: 'Oui, en lecture seule : les valeurs, les formules lorsqu’elles se décodent, la mise en forme de base, les fusions et la disposition des feuilles se chargent, et le classeur s’enregistre en .xlsx.',
			},
			{
				q: 'Et les fichiers protégés par mot de passe ?',
				a: 'Non pris en charge aujourd’hui.',
			},
		],
	},
	finale: {
		kicker: 'Commencer',
		title: '.xlsx en entrée. .xlsx en sortie.',
		sub: 'Lisez le guide de prise en main, choisissez votre framework et essayez la démo avec un classeur de votre choix. Sous licence Apache-2.0, TypeScript strict, et honnête sur les écarts.',
		quick: { text: 'Lire le guide', href: '/fr/getting-started' },
		github: 'Voir sur GitHub',
		columns: [
			{
				title: 'Produit',
				links: [
					{
						text: 'Démo en direct',
						href: 'https://christophervr.github.io/ooxml/xlsx/demo/',
						external: true,
					},
					{ text: 'Liaisons de frameworks', href: '/bindings' },
					{ text: 'API de l’élément', href: '/fr/api' },
					{ text: 'Versions', href: '/releasing' },
				],
			},
			{
				title: 'Documentation',
				links: [
					{ text: 'Prise en main', href: '/fr/getting-started' },
					{ text: 'Démos en direct', href: '/demos' },
					{ text: 'Architecture', href: '/fr/architecture' },
					{ text: 'Thèmes', href: '/theming' },
					{ text: 'Fonctionnalités et limites', href: '/fr/features' },
				],
			},
			{
				title: 'Communauté',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx',
						external: true,
					},
					{
						text: 'Tickets',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Licence',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'xlsx-viewer · un éditeur de tableur précoce pour le web',
	},
};
