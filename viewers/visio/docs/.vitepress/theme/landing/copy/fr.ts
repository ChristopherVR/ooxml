import type { LandingCopy } from './types';

export const fr: LandingCopy = {
	hero: {
		kicker: 'Bêta publique · Apache-2.0 · TypeScript',
		titleTop: 'Affichage de .vsdx,',
		titleAccent: 'conçu pour être intégré.',
		sub: 'Intégrez des diagrammes Visio à votre application web. Un seul composant web <visio-viewer> partagé et un seul moteur de rendu SVG pour React, Vue, Angular, Svelte, Solid et JavaScript vanilla. Parcourez les pages, inspectez les formes et effectuez les modifications de texte brut prises en charge ; les adaptateurs de frameworks ne font que relier le cycle de vie et les événements.',
		start: { text: 'Commencer', href: '/fr/getting-started' },
		demo: 'Démo en direct',
		scroll: 'Défiler',
		frameCaption: 'Exemple de flux de travail · en direct dans le navigateur',
		frameTry: 'Essayer',
		frameAlt:
			"L'espace de travail visio-viewer affichant l'exemple original, la navigation entre les pages, le ruban et l'inspecteur de formes",
		copyLabel: 'Copier',
		copiedLabel: 'Copié',
		entryLabel: 'installer',
		notPublished:
			'Publié sur npm. Chaque paquet de framework est autonome : il regroupe le visualiseur partagé, son adaptateur et ses workers.',
	},
	features: {
		kicker: 'Fonctionnalités',
		title: "Ce qui existe aujourd'hui.",
		items: [
			{
				title: 'Six liaisons natives',
				copy: 'React, Vue, Angular, Svelte, Solid ou JavaScript vanilla. De fins adaptateurs de cycle de vie laissent toute la logique de format dans ooxml-core et le visualiseur partagé.',
				link: { text: 'Guides par framework', href: '/frameworks/react' },
			},
			{
				title: 'Un diagramme au premier plan',
				copy: 'Changez de page, ajustez et zoomez, sélectionnez des formes, recherchez le texte visible, réglez la visibilité des calques prise en charge, et utilisez le ruban de style Visio, le menu Fichier, la fenêtre Formes et la fenêtre Panoramique et zoom.',
				link: { text: 'Démos en direct', href: '/demos' },
			},
			{
				title: 'Votre navigateur, vos fichiers',
				copy: 'Ouvrez un fichier .vsdx local et exportez la page courante en SVG. Aucun envoi de document, aucune télémétrie et aucun chargement de document distant.',
				link: { text: 'Prise en main', href: '/fr/getting-started' },
			},
			{
				title: 'Une édition prudente',
				copy: "Remplacement expérimental de texte brut adossé à la source, modifications de géométrie ciblées, annulation/rétablissement bornés et copies VSDX explicites, d'origine ou modifiées. Le contenu non pris en charge est refusé, pas supprimé.",
				link: { text: "Limites de l'édition", href: '/fr/getting-started#editing' },
			},
			{
				title: 'Aperçu des anciens .vsd',
				copy: "Les fichiers VSD binaires v11 se chargent via le codec ole2 partagé d'ooxml-core, en aperçu en lecture seule avec des styles de repli diagnostiqués.",
				link: { text: 'Registre des capacités', href: '/parity' },
			},
			{
				title: 'Partage en direct',
				copy: "Fichier > Partager démarre une session Yjs via BroadcastChannel entre des fenêtres d'un même navigateur. Le partage entre appareils nécessite un transport serveur fourni par l'hôte, qui n'est pas réalisé.",
				link: { text: 'Collaboration', href: '/fr/collaboration' },
			},
			{
				title: "Export et préparation à l'impression",
				copy: "Export SVG de la page courante avec images raster intégrées et notes de compatibilité visibles, ainsi que des instantanés de pages immuables via l'API. Il n'y a ni boîte de dialogue d'impression, ni export PDF ou bitmap.",
				link: { text: 'API du visualiseur', href: '/fr/api' },
			},
			{
				title: 'Limites visibles',
				copy: "Chaque capacité est suivie avec ses preuves et ce qui manque encore. Un beau ruban n'est pas présenté comme une parité avec Visio.",
				link: { text: 'Registre des capacités', href: '/parity' },
			},
		],
	},
	status: {
		kicker: 'État',
		title: 'Honnête sur ce qui reste à faire.',
		copy: "Il s'agit d'une bêta publique, pas d'une parité avec Microsoft Visio. Le dessin général, le recalcul complet de ShapeSheet, le texte enrichi, les groupes et la plupart des masques, l'export PDF et la fidélité de réouverture dans Visio natif ne sont pas établis. L'exemple généré illustre les interactions ; il ne prouve pas la fidélité visuelle.",
		link: { text: 'Registre des capacités', href: '/parity' },
		roadmapTitle: 'Où se trouve la logique',
		roadmapCopy:
			"L'analyse, l'interprétation de ShapeSheet, la géométrie et les commandes d'édition résident dans ooxml-core/visio. Ce dépôt ne contient que l'interface : un contrôleur, un moteur de rendu SVG, un élément et six fins adaptateurs.",
		roadmapLink: { text: 'Architecture', href: '/fr/architecture' },
	},
	quickstart: {
		kicker: 'Prise en main',
		title: 'Montez un visualiseur en quelques lignes.',
		copy: 'Chaque adaptateur transmet les mêmes propriétés, les mêmes événements et le même handle impératif du visualiseur partagé. Donnez une hauteur au conteneur, puis passez un VisioDocument ou chargez un File.',
		docsLabel: 'Guide du framework',
		buildTitle: 'Exécuter depuis les sources',
		buildCopy:
			'Pour lancer le terrain de jeu ou travailler sur le visualiseur lui-même, clonez le dépôt. Node.js 22.12 ou plus récent est requis.',
		buildCommands: 'npm ci --ignore-scripts\nnpm run dev',
	},
	demos: {
		kicker: 'Démo en direct',
		title: 'Essayez-le ici même.',
		copy: "Voici le vrai visualiseur exécuté dans votre navigateur : le terrain de jeu construit à partir de ce dépôt pour chaque adaptateur de framework, intégré en direct. Ouvrez le dessin d'exemple ou un .vsdx local ; rien n'est envoyé.",
		frameworkLabel: 'Framework',
		soloTab: 'Visualiseur',
		collabTab: 'Partage',
		guestPicker: 'Framework de la fenêtre B',
		load: 'Charger la démo en direct',
		loading: 'Chargement du visualiseur en direct',
		openFull: "Ouvrir l'application complète",
		hostLabel: 'Fenêtre A',
		guestLabel: 'Fenêtre B',
		soloHint:
			"Utilisez le dessin d'exemple ou ouvrez un fichier .vsdx de votre choix. Changer de framework démarre un nouveau visualiseur.",
		collabHint:
			"Deux fenêtres d'un même navigateur, chacune dans le framework de votre choix, ont déjà rejoint par son nom une même session Fichier > Partager. La fenêtre A ouvre l'exemple et la fenêtre B le reçoit. Modifiez le texte d'une forme dans l'une et il apparaît dans l'autre. Elle utilise un BroadcastChannel identifié par le nom de la session ; il n'y a pas de serveur, cela ne fonctionne qu'à l'intérieur de ce navigateur, et rien n'est envoyé.",
	},
	faq: {
		kicker: 'FAQ',
		title: 'Questions fréquentes.',
		items: [
			{
				q: 'Est-ce un remplacement complet de Visio ?',
				a: "Non. Il s'agit d'une implémentation précoce. La géométrie en cache, le texte et les images pris en charge sont rendus par un seul chemin SVG. Les documents complexes présentent encore des lacunes ; le registre des capacités distingue l'implémentation, les preuves et le travail restant.",
				link: { text: 'Registre des capacités', href: '/parity' },
			},
			{
				q: "Puis-je l'installer depuis npm ?",
				a: 'Oui. Un paquet de visualiseur autonome par framework est disponible sur npm, par exemple npm install visio-react-viewer. Chaque paquet est versionné séparément.',
				link: { text: 'Prise en main', href: '/fr/getting-started' },
			},
			{
				q: "L'ouverture d'un fichier l'envoie-t-elle quelque part ?",
				a: "Non. Le visualiseur lit le File que vous sélectionnez dans votre navigateur. Il n'y a ni envoi de document, ni télémétrie, ni récupération de document distant.",
			},
			{
				q: 'Puis-je modifier et enregistrer un fichier .vsdx ?',
				a: "L'édition expérimentale du texte brut et de la géométrie ciblée est disponible pour les formes prises en charge adossées à la source, avec annulation/rétablissement bornés et téléchargements explicites de copies VSDX. Le texte lié à un masque, le texte enrichi, les champs, les paquets signés et le contenu avec macros sont refusés. Le fichier source n'est jamais écrasé, et la réouverture dans Visio natif reste non vérifiée.",
				link: { text: 'API du visualiseur', href: '/fr/api' },
			},
			{
				q: 'Que puis-je exporter ou imprimer ?',
				a: "L'interface exporte la page courante en SVG ainsi que des copies VSDX d'origine ou modifiées. L'API peut préparer des instantanés de pages immuables pour un futur flux d'impression ; elle n'ouvre pas de boîte de dialogue d'impression. L'export PDF, l'export bitmap et la mise en page d'imprimante native ne sont pas implémentés.",
			},
			{
				q: 'Plusieurs personnes peuvent-elles modifier un même diagramme ?',
				a: "Seulement de façon limitée : Fichier > Partager relie des fenêtres d'un même navigateur par une session Yjs. Le partage entre appareils nécessite un transport serveur fourni par l'hôte, que cette version n'inclut pas.",
				link: { text: 'Collaboration', href: '/fr/collaboration' },
			},
			{
				q: "L'exemple prouve-t-il la fidélité visuelle ?",
				a: "Non. L'exemple original, créé par programme, illustre les interactions. Les fixtures générées n'établissent pas la parité visuelle avec Microsoft Visio ; comparez les dessins importants avec leur rendu dans Visio natif.",
			},
			{
				q: 'Quels frameworks sont pris en charge ?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS et JavaScript pur. Ils partagent un même élément de visualiseur ; chaque adaptateur ne gère que le montage, les mises à jour des propriétés et la transmission des événements.',
				link: { text: 'Guides par framework', href: '/frameworks/react' },
			},
		],
	},
	finale: {
		kicker: 'Commencer',
		title: '.vsdx en entrée. Un diagramme en sortie.',
		sub: 'Lisez le guide, choisissez votre framework et essayez le terrain de jeu avec un dessin de votre choix. Sous licence Apache-2.0, TypeScript strict, et honnête sur les lacunes.',
		quick: { text: 'Lire le guide', href: '/fr/getting-started' },
		github: 'Voir sur GitHub',
		columns: [
			{
				title: 'Produit',
				links: [
					{
						text: 'Démo en direct',
						href: 'https://christophervr.github.io/ooxml/visio/demo/',
						external: true,
					},
					{ text: 'Démos en direct', href: '/demos' },
					{ text: 'Collaboration', href: '/fr/collaboration' },
					{ text: 'Versions', href: '/releasing' },
				],
			},
			{
				title: 'Documentation',
				links: [
					{ text: 'Prise en main', href: '/fr/getting-started' },
					{ text: 'Architecture', href: '/fr/architecture' },
					{ text: 'API du visualiseur', href: '/fr/api' },
					{ text: 'Registre des capacités', href: '/parity' },
				],
			},
			{
				title: 'Communauté',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio',
						external: true,
					},
					{
						text: 'Tickets',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Licence',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'visio-viewer · un visualiseur Visio précoce pour le web',
	},
};
