import type { LandingCopy } from './types';

export const fr: LandingCopy = {
	hero: {
		titleTop: 'Discussions et appels d’équipe,',
		titleAccent: 'sur un serveur que vous gérez.',
		sub: 'OpenTeams est un espace de travail d’équipe pour le web : canaux et discussions, présence, réunions WebRTC et fichiers Office partagés dans la conversation. Un seul composant web, de fines intégrations pour six frameworks et un petit serveur de référence. Une implémentation précoce : ce n’est pas Microsoft Teams, et il n’y a pas de chiffrement de bout en bout.',
		start: { text: 'Commencer', href: '/fr/getting-started' },
		demo: { text: 'Essayer la démo', href: '/demo/' },
		frameLabel: '<teams-app> en mode local',
		tryLabel: 'Ouvrir la démo',
		imageAlt:
			'La démo OpenTeams : une liste de canaux, une conversation avec des réponses et des réactions, et la zone de saisie',
	},
	liveDemo: {
		title: 'Essayez-la ici même',
		note: 'Démo en direct',
		introBefore:
			'Voici la vraie application, exécutée dans votre navigateur, sans serveur derrière : discussions, présence et appels circulent uniquement entre les cadres et les onglets de ce navigateur (BroadcastChannel). Passez à deux personnes pour réunir Ada et Bob dans une même salle, chacun dans le framework de votre choix (par exemple React pour Ada et Vue pour Bob), et regardez un message passer d’un volet à l’autre. Un usage réel nécessite un serveur ; voir ',
		introLink: { text: 'apportez votre propre serveur', href: '/fr/server' },
		introAfter: '.',
		frameworkLabel: 'Framework d’Ada',
		soloTab: 'Une personne',
		duoTab: 'Deux personnes',
		guestPicker: 'Framework de Bob',
		load: 'Charger la démo en direct',
		openFull: 'Ouvrir l’application complète',
		frameTitle: 'Démo OpenTeams {framework}',
		hintDuo:
			'Deux applications distinctes, chacune construite avec sa propre intégration de framework, rejoignent une même salle par son nom et la partagent via BroadcastChannel. Le framework importe peu : elles parlent le même protocole du cœur. Il s’agit du mode local du cœur, pas d’un service hébergé : avec votre propre serveur, les deux mêmes personnes peuvent se trouver sur des machines différentes.',
		hintSolo:
			'Tout s’exécute côté client et reste dans ce navigateur. Ouvrez l’application complète, ou passez à deux personnes pour voir la présence et la discussion entre deux utilisateurs.',
	},
	embed: {
		title: 'Deux façons de l’intégrer',
		note: 'Applications de démo',
		openDemo: 'Ouvrir la démo',
		demos: [
			{
				title: 'Vanilla : l’application complète',
				copy: 'Un seul élément <teams-app> : canaux, messages avec réponses et réactions, fichiers, recherche, présence et réunions. Ouvrez-le dans deux onglets pour voir deux personnes discuter.',
				href: '/demo/',
				link: '',
			},
			{
				title: 'React : composant et hook',
				copy: 'Le composant <Teams /> à gauche et un panneau écrit à la main au-dessus du hook useTeams() à droite, tous deux sur un même espace de travail.',
				href: '/demo-react/',
				link: '',
			},
		],
	},
	features: {
		title: 'Un espace de travail que vous hébergez',
		note: 'Ce qu’il fait',
		items: [
			{
				title: 'Canaux et discussions',
				copy: 'Canaux, messages avec réponses, réactions, modification et suppression réservées à l’auteur, recherche et compteurs de non-lus, le tout dans un seul document Yjs qui fusionne les modifications concurrentes.',
				href: '/fr/getting-started',
				link: 'Premiers pas',
				icon: 'chat',
			},
			{
				title: 'Présence',
				copy: 'Qui est en ligne, indicateurs de saisie et une disponibilité que vous définissez vous-même (disponible, occupé, absent). La présence est indicative, comme toute awareness Yjs.',
				href: '/architecture',
				link: 'Architecture',
				icon: 'people',
			},
			{
				title: 'Réunions via WebRTC',
				copy: 'Audio, vidéo, partage d’écran et main levée dans un appel en maillage complet : pas de serveur média, honnête jusqu’à une douzaine de personnes. Le TURN, c’est à vous de l’exploiter.',
				href: '/fr/server#appels-stun-et-turn',
				link: 'Appels et TURN',
				icon: 'video',
			},
			{
				title: 'Des fichiers dans la conversation',
				copy: 'Les pièces jointes sont envoyées sur votre serveur (ou votre propre service de dépôt) et publiées sous forme de liens. Les fichiers Word, Excel, PowerPoint et Visio sont reconnus afin qu’un hôte puisse les ouvrir.',
				href: '/fr/server#fichiers',
				link: 'Stockage des fichiers',
				icon: 'file',
			},
			{
				title: 'Apportez votre propre serveur',
				copy: 'Trois petits contrats : synchronisation y-websocket, un relais de signalisation JSON et un stockage de fichiers facultatif. openteams-server fait les trois dans un seul processus Node.',
				href: '/fr/server',
				link: 'Guide du serveur',
				icon: 'server',
			},
			{
				title: 'Six frameworks',
				copy: 'Intégrations React, Vue, Angular, Svelte, Solid et JS natif. Chacune fournit le composant complet et un hook brut au-dessus du même store client.',
				href: '/frameworks/react',
				link: 'Guides par framework',
				icon: 'code',
			},
		],
	},
	limits: {
		title: 'Ce qu’il n’est pas',
		note: 'Des limites assumées',
		items: [
			'Pas de chiffrement de bout en bout : votre serveur stocke et peut lire chaque message et chaque fichier.',
			'L’autorisation est l’affaire du serveur ; le serveur de référence n’a qu’un jeton partagé, pas de comptes utilisateur.',
			'Les appels en maillage conviennent à une poignée de personnes, pas à un webinaire.',
			'Pas de discussions 1:1, de notifications, de calendrier, d’enregistrement, de flou d’arrière-plan ni de sous-titres en direct.',
			'Ce n’est pas Microsoft Teams, il n’est pas affilié à Microsoft et ne prétend à aucune équivalence fonctionnelle.',
		],
		all: { text: 'Toutes les limites', href: '/fr/limitations' },
	},
	quickstart: {
		title: 'Un espace de travail en quelques lignes',
		note: 'Démarrage rapide',
		introBefore:
			'Installez l’intégration de votre framework et affichez le composant. Le mode local n’a pas besoin de serveur ; pour un usage réel, faites pointer ',
		introAfter:
			' vers vos propres points d’accès de synchronisation et de signalisation (le serveur de référence fournit les deux). Angular, Svelte et Solid offrent les mêmes options.',
		fullGuide: 'Guide complet : {framework}',
	},
	faq: {
		title: 'Questions fréquentes',
		note: 'FAQ',
		items: [
			{
				q: 'Peut-on l’utiliser gratuitement à titre commercial ?',
				a: 'Le dépôt est sous licence Apache-2.0. Les paquets sont publiés sur npm sous les noms openteams-*-viewer et openteams-server.',
			},
			{
				q: 'A-t-il besoin d’un serveur ?',
				a: 'Pour un usage réel, oui : vous apportez votre propre serveur pour la synchronisation (y-websocket), la signalisation et, en option, le stockage de fichiers. Le serveur de référence, openteams-server, fait les trois dans un seul processus Node. Sans serveur, les démos fonctionnent dans un mode local qui ne relie que les onglets d’un même navigateur.',
				href: '/fr/server',
				link: 'Apportez votre propre serveur',
			},
			{
				q: 'Est-il chiffré de bout en bout ?',
				a: 'Non. Les discussions transitent sous forme de mises à jour Yjs par votre serveur, qui peut les lire et les stocker. Le flux média des appels est en DTLS-SRTP entre pairs.',
				href: '/fr/limitations',
				link: 'Limites',
			},
			{
				q: 'Est-ce Microsoft Teams ?',
				a: 'Non. OpenTeams n’est pas Microsoft Teams, n’est pas affilié à Microsoft et ne prétend à aucune équivalence fonctionnelle. Il offre des canaux, des messages avec réponses, des réactions, des fichiers, la recherche, la présence et des réunions avec partage d’écran.',
			},
			{
				q: 'Combien de personnes peuvent rejoindre un appel ?',
				a: 'Les appels sont en maillage complet, sans serveur média : ils conviennent à une poignée de personnes (plafond par défaut de 12), pas à un webinaire. Le TURN, c’est à vous de l’exploiter.',
				href: '/fr/server#appels-stun-et-turn',
				link: 'Appels et TURN',
			},
			{
				q: 'Quels frameworks sont pris en charge ?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS et JavaScript natif. Ils partagent un seul élément <teams-app> et un seul client du cœur ; chaque intégration expose aussi un hook brut pour construire votre propre interface.',
				href: '/frameworks/react',
				link: 'Guides par framework',
			},
			{
				q: 'Fonctionne-t-il avec les fichiers Word, Excel, PowerPoint et Visio ?',
				a: 'Les fichiers sont partagés sous forme de liens. Les fichiers Office sont reconnus par leur nom afin qu’une application hôte puisse les ouvrir dans la visionneuse correspondante ; OpenTeams ne les modifie pas lui-même.',
			},
		],
	},
	finale: {
		title: 'Votre espace de travail, votre serveur.',
		introBefore: 'Installez une intégration, démarrez le serveur de référence avec ',
		introAfter:
			', et essayez la démo avec deux onglets. Sous licence Apache-2.0, TypeScript strict, et honnête sur les lacunes.',
		start: { text: 'Commencer', href: '/fr/getting-started' },
		github: 'Voir sur GitHub',
		columns: [
			{
				title: 'Produit',
				links: [
					{ text: 'Démo en direct', href: '/demo/', external: false },
					{ text: 'Apportez votre propre serveur', href: '/fr/server', external: false },
					{ text: 'Collaboration', href: '/fr/collaboration', external: false },
					{ text: 'Versions', href: '/releasing', external: false },
				],
			},
			{
				title: 'Documentation',
				links: [
					{ text: 'Premiers pas', href: '/fr/getting-started', external: false },
					{ text: 'Architecture', href: '/architecture', external: false },
					{ text: 'API de l’élément', href: '/fr/api', external: false },
					{ text: 'Limites', href: '/fr/limitations', external: false },
				],
			},
			{
				title: 'Communauté',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams',
						external: true,
					},
					{
						text: 'Tickets',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Licence',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/LICENSE',
						external: true,
					},
				],
			},
		],
		license: '© 2026 ChristopherVR · Apache-2.0',
		disclaimer: 'OpenTeams · pas Microsoft Teams, non affilié à Microsoft',
	},
};
