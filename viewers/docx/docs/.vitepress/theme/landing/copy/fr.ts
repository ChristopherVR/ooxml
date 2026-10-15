import type { LandingCopy } from './types';

export const fr: LandingCopy = {
	hero: {
		kicker: 'Implémentation précoce · Apache-2.0 · TypeScript',
		titleTop: "L'édition .docx,",
		titleAccent: 'prête à intégrer.',
		sub: 'Un éditeur Word dans le navigateur pour React, Vue, Angular, Svelte, Solid et JavaScript vanilla. Un seul modèle de document et un composant web <docx-editor> reposant sur ProseMirror font le travail ; les adaptateurs de frameworks ne gèrent que le cycle de vie et les événements.',
		start: { text: 'Commencer', href: '/fr/getting-started' },
		demo: 'Démo en direct',
		scroll: 'Défiler',
		frameCaption: 'Document d’exemple.docx · en direct dans le navigateur',
		frameTry: 'Essayer',
		frameAlt: 'L’éditeur docx-viewer affichant un document d’exemple avec le ruban de style Word',
		copyLabel: 'Copier',
		copiedLabel: 'Copié',
		entryLabel: 'installer',
		notPublished:
			'Publié sur npm. Chaque paquet de framework est autonome : il embarque l’éditeur et apporte docx-core avec lui.',
	},
	features: {
		kicker: 'Fonctionnalités',
		title: 'Ce qui existe aujourd’hui.',
		items: [
			{
				title: 'Un seul éditeur partagé',
				copy: 'Un ruban de style Word, une barre de titre, la vue Fichier et une barre d’état vivent dans un seul composant web, avec annulation/rétablissement, recherche et remplacement, zoom et un mode lecture seule. Les adaptateurs n’ajoutent aucune interface propre.',
				link: { text: 'Architecture', href: '/fr/architecture' },
			},
			{
				title: 'Mise en page Impression',
				copy: 'Une vue paginée avec coupure de lignes, règles de veuves et orphelines et de lignes solidaires, colonnes, scission des lignes de tableau, en-têtes répétés, placement des notes de bas de page et impression. C’est l’approximation propre à cet éditeur, pas la pagination de Word.',
				link: { text: 'Feuille de route de la prise en charge', href: '/parity-roadmap' },
			},
			{
				title: 'Structure du document',
				copy: 'Styles, listes numérotées et à puces, tableaux, sections et colonnes, en-têtes et pieds de page, notes de bas de page et champs, dont la table des matières. Des zones de texte en ligne, un filigrane d’en-tête et des bordures de page peuvent être insérés.',
				link: { text: 'Travaux restants', href: '/outstanding-work' },
			},
			{
				title: 'Modifications suivies et commentaires',
				copy: 'Mode Suivi des modifications, acceptation et rejet, révision suivante et précédente, et un volet de commentaires avec réponses et résolution. Les commentaires sont réécrits dans comments.xml.',
				link: { text: 'Édition du texte', href: '/fr/editing' },
			},
			{
				title: 'Préservation rigoureuse',
				copy: 'Un enregistrement sans changement renvoie les octets d’origine. Les modifications prises en charge préservent les parties du paquet extérieures au modèle, et les modifications qui endommageraient un contenu non pris en charge sont rejetées au lieu d’être perdues.',
				link: { text: 'Modèle et enregistrement', href: '/fr/architecture' },
			},
			{
				title: 'Ancien format .doc',
				copy: 'Les fichiers Word 97-2003 se chargent grâce aux codecs ole2 partagés. Le texte du corps principal est importé, et des modifications limitées de paragraphes existants sont réenregistrées en .doc. Les changements de mise en forme et de structure sont rejetés.',
				link: { text: 'Feuille de route de la prise en charge', href: '/parity-roadmap' },
			},
			{
				title: 'Protocole de collaboration',
				copy: 'Un protocole neutre vis-à-vis du transport échange des étapes ProseMirror validées par l’intermédiaire d’une autorité, avec une présence transitoire. Votre application gère le réseau, l’identité, les permissions et le stockage.',
				link: { text: 'Collaboration', href: '/fr/collaboration' },
			},
			{
				title: 'Localisation',
				copy: 'L’interface de l’éditeur est livrée en anglais, français, allemand, espagnol et chinois simplifié via une seule option de locale. Elle ne traduit que le texte de l’interface, jamais le contenu du document.',
				link: { text: 'Adaptateurs', href: '/bindings' },
			},
		],
	},
	status: {
		kicker: 'État',
		title: 'Honnête sur ce qui n’est pas terminé.',
		copy: 'Il s’agit d’une implémentation précoce. Elle n’offre pas la parité avec Microsoft Word : la pagination est approximative, de nombreuses fonctionnalités ne sont pas rendues avec la fidélité de Word, les fichiers protégés par mot de passe ne sont pas pris en charge, et l’export DOCX ne peut pas récupérer le contenu que le modèle omet. Les pages de feuille de route et de travaux restants détaillent les lacunes.',
		link: { text: 'Feuille de route de la prise en charge', href: '/parity-roadmap' },
		roadmapTitle: 'Note de feuille de route',
		roadmapCopy:
			'La logique OOXML partagée est en cours de déplacement vers le paquet privé ooxml-core, afin que ce dépôt puisse ne contenir plus que l’interface. Cette migration est en cours et n’est pas une fonctionnalité livrée.',
		roadmapLink: { text: 'Le plan d’ooxml-core', href: '/ooxml-core-plan' },
	},
	quickstart: {
		kicker: 'Prise en main',
		title: 'Montez un éditeur en quelques lignes.',
		copy: 'Chaque adaptateur monte le même <docx-editor>. Transmettez un modèle de document, écoutez les changements et donnez une hauteur au conteneur.',
		docsLabel: 'Guide du framework',
		buildTitle: 'Construire depuis les sources',
		buildCopy:
			'Pour exécuter les démos ou travailler sur l’éditeur lui-même, clonez le dépôt. Bun est requis.',
		buildCommands: 'bun install\nbun run demo',
	},
	demos: {
		kicker: 'Démo en direct',
		title: 'Essayez-le ici même.',
		copy: 'Voici le véritable éditeur exécuté dans votre navigateur : l’application de démo construite à partir de ce dépôt pour chaque adaptateur de framework, intégrée en direct. Ouvrez le document d’exemple, ou scindez la vue pour voir deux pairs locaux partager un même document.',
		frameworkLabel: 'Framework',
		languageLabel: 'Langue',
		soloTab: 'Éditeur',
		collabTab: 'Collaboration',
		guestPicker: 'Framework de la fenêtre B',
		load: 'Charger la démo en direct',
		loading: 'Chargement de l’éditeur en direct',
		openFull: 'Ouvrir l’application complète',
		hostLabel: 'Hôte (ouvre l’exemple)',
		guestLabel: 'Rejoint par le nom de session',
		soloHint:
			'Utilisez « Open the sample document » dans la démo, ou déposez un fichier DOCX ou DOC de votre choix. Changer de framework démarre un nouvel éditeur.',
		collabHint:
			'Deux fenêtres, chacune dans le framework de votre choix, rejoignent une même session par son nom. La fenêtre hôte exécute l’autorité de référence en mémoire ; l’autre reçoit le document et chaque modification, avec les curseurs de chacun. La session est un BroadcastChannel : elle ne fonctionne qu’entre des fenêtres de ce navigateur ; il n’y a pas de serveur réseau, et le transport de production est l’affaire de l’application hôte.',
	},
	faq: {
		kicker: 'FAQ',
		title: 'Questions fréquentes.',
		items: [
			{
				q: 'Puis-je l’installer depuis npm ?',
				a: 'Oui. docx-core et un paquet d’éditeur autonome par framework sont sur npm, par exemple npm install docx-react-viewer. Chaque paquet est versionné séparément.',
				link: { text: 'Politique de publication', href: '/releasing' },
			},
			{
				q: 'L’usage commercial est-il gratuit ?',
				a: 'Le dépôt est sous licence Apache-2.0.',
			},
			{
				q: 'Est-ce la parité avec Word ?',
				a: 'Non. La pagination est une approximation, de nombreuses fonctionnalités ne sont pas rendues avec la fidélité de Word, et l’export DOCX crée un nouveau document à partir du contenu visible pris en charge. La parité est suivie séparément pour l’import, la mise en page, l’édition, l’export, la préservation et l’accessibilité.',
				link: { text: 'Feuille de route de la prise en charge', href: '/parity-roadmap' },
			},
			{
				q: 'Les fichiers modifiés se rouvrent-ils dans Word ?',
				a: 'Les modifications prises en charge sont réécrites dans le paquet d’origine, et les modifications risquées sont rejetées au lieu de le dégrader silencieusement. C’est un objectif de conception vérifié par des tests de régression, pas une promesse d’export sans perte.',
				link: { text: 'Architecture', href: '/fr/architecture' },
			},
			{
				q: 'Quels frameworks sont pris en charge ?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS et JavaScript pur. Ils partagent un seul éditeur et un seul modèle ; chaque adaptateur ne gère que le montage, les mises à jour de propriétés et la transmission des événements.',
				link: { text: 'Adaptateurs', href: '/bindings' },
			},
			{
				q: 'La collaboration nécessite-t-elle un serveur ?',
				a: 'Le protocole est neutre vis-à-vis du transport et fournit une autorité de référence en mémoire. Le réseau, l’authentification, la persistance et les permissions relèvent de la responsabilité de votre application.',
				link: { text: 'Collaboration', href: '/fr/collaboration' },
			},
			{
				q: 'Peut-il ouvrir d’anciens fichiers .doc ?',
				a: 'Oui, avec des limites. Le texte du corps principal est importé et des modifications limitées de paragraphes existants sont réenregistrées en .doc ; les changements de mise en forme et de structure sont rejetés.',
			},
			{
				q: 'Et les fichiers protégés par mot de passe ?',
				a: 'Non pris en charge aujourd’hui.',
			},
		],
	},
	finale: {
		kicker: 'Commencer',
		title: '.docx en entrée. .docx en sortie.',
		sub: 'Lisez l’architecture, choisissez le guide de votre framework et essayez la démo avec un document de votre choix. Sous licence Apache-2.0, en TypeScript strict, et honnête sur les lacunes.',
		quick: { text: 'Lire le guide', href: '/fr/getting-started' },
		github: 'Voir sur GitHub',
		columns: [
			{
				title: 'Produit',
				links: [
					{
						text: 'Démo en direct',
						href: 'https://christophervr.github.io/ooxml/docx/demo/',
						external: true,
					},
					{ text: 'Adaptateurs de frameworks', href: '/bindings' },
					{ text: 'Collaboration', href: '/fr/collaboration' },
					{ text: 'Versions', href: '/releasing' },
				],
			},
			{
				title: 'Documentation',
				links: [
					{ text: 'Prise en main', href: '/fr/getting-started' },
					{ text: 'Architecture', href: '/fr/architecture' },
					{ text: 'Édition du texte', href: '/fr/editing' },
					{ text: 'Feuille de route de la prise en charge', href: '/parity-roadmap' },
					{ text: 'Travaux restants', href: '/outstanding-work' },
				],
			},
			{
				title: 'Communauté',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx',
						external: true,
					},
					{
						text: 'Tickets',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Licence',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'docx-viewer · un éditeur Word précoce pour le web',
	},
};
