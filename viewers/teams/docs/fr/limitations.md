# Limites

OpenTeams est en bêta. Cette page énumère ce qu’il ne fait pas, afin que vous
puissiez décider avant de vous y fier.

## Pas Microsoft Teams

OpenTeams n’est pas Microsoft Teams, n’est pas affilié à Microsoft et ne prétend à aucune
équivalence fonctionnelle. Il implémente :

- des canaux, des messages avec réponses et volets de fils, des réactions, la modification et la suppression
  réservées à l’auteur, la recherche et des compteurs de non-lus ;
- des discussions 1:1 et de groupe, chacune dans sa propre salle que seuls ses membres peuvent rejoindre ;
- un fil d’activité et des notifications dans l’application (message, notification du navigateur tant que l’onglet est ouvert, son) ;
- un calendrier avec réunions, répétitions, rappels et export `.ics` ;
- la présence : qui est en ligne, qui écrit, et un statut que vous définissez (disponible, occupé, ne pas déranger, de retour bientôt, absent, hors ligne) avec une heure de réinitialisation facultative, un message de statut et le passage automatique en absent en cas d’inactivité ;
- des réunions dans un canal : audio, vidéo, partage d’écran et main levée, avec un écran de
  pré-connexion ;
- des fichiers partagés dans la conversation sous forme de liens, les fichiers Word, Excel,
  PowerPoint et Visio étant reconnus par leur nom.

Il n’implémente **pas** l’enregistrement, le flou d’arrière-plan, les sous-titres en direct, les notifications push ou par e-mail
(les notifications ne fonctionnent que lorsque l’application est ouverte), les invitations hors de l’espace de travail ni les réponses (RSVP).

## Espace de contenu

Word s’affiche en lecture seule ; Excel permet l’édition locale et l’enregistrement de nouvelles
copies dans le canal lorsque le stockage est configuré ; Visio utilise sa visionneuse existante.
Il n’y a ni co-édition, ni écrasement, ni réenregistrement versionné. PowerPoint nécessite une page
de visionneuse intégrée configurée par l’hôte. Les intégrations manquantes sont signalées dans le
volet. La lecture de texte est limitée à 2 Mio et celle des fichiers Office natifs à 32 Mio. L’interface existe en anglais, allemand, espagnol, français et chinois simplifié (propriété `locale`) ; le texte écrit par les personnes n’est jamais traduit.

Les vues Fichiers permettent l’envoi direct, la création d’un classeur Excel vierge et la recherche
dans les métadonnées lorsque le stockage est configuré. Les envois affichent un état d’occupation et
permettent de réessayer, sans progression en octets ni annulation. Un lot en échec peut laisser des fichiers envoyés avant l'échec que aucun message ne référence ; le serveur de
référence supprime ses envois partiels et laisse la personne qui a envoyé le fichier ou un modérateur supprimer le reste
(`DELETE /files/<scope>/<name>`), mais ne le fait pas de lui-même. Le serveur de référence limite les fichiers des chats
privés à leurs membres et compte le stockage par espace de travail ; les fichiers de canal sont lisibles par tout
utilisateur connecté du serveur. Les dossiers, l'indexation du contenu des documents et l'historique des versions de
fichiers ne sont pas implémentés.

Markdown prend en charge les titres, les puces, les citations, le code délimité, l’emphase de base,
les liens web, les listes de tâches plates, les tableaux de premier niveau et les liens web relatifs au fichier. Il n’implémente pas CommonMark/GFM complet,
les structures imbriquées ni les liens de référence. Le HTML brut est affiché comme du texte.

Les sites web et les pièces jointes HTML utilisent un bac à sable sans permission de même origine,
de fenêtre contextuelle ni de navigation de niveau supérieur. Certains sites bloquent l’intégration
ou exigent des permissions que ce bac à sable n’accorde pas. L’ouverture externe reste disponible.
Les définitions d’onglets de fichiers et de sites web peuvent être partagées dans les canaux ; la
plateforme d’applications intégrées, l’authentification et les permissions de Microsoft Teams ne
sont pas implémentées. Les contrôles du créateur d’un onglet s’exécutent dans le client et exigent
une autorisation côté serveur pour devenir applicables.

## Sécurité et confidentialité

- **Pas de chiffrement de bout en bout.** Les discussions transitent sous forme de mises à jour Yjs
  par votre serveur, qui peut les lire et les stocke. Les fichiers envoyés sont stockés tels
  quels. Le flux média des appels est en DTLS-SRTP entre pairs (un relais TURN ne peut pas le
  déchiffrer).
- **L’autorisation est l’affaire du serveur.** Les contrôles d’auteur et de rôle s’exécutent dans
  le client ; un client hostile peut écrire tout ce que le serveur accepte.
- **L’authentification est obligatoire par défaut.** Le serveur de référence refuse de démarrer sans comptes, sans fonction `authenticate` ni `authorizeSync` ; `--insecure-dev` active un mode de développement ouvert, avec un avertissement. Avec des comptes, il signe des sessions pour des utilisateurs configurés (ou votre fonction `authenticate`) et prend l’identifiant d’utilisateur dans la session, et `authorizeSync` peut décider de l’accès par salle. Un `TEAMS_TOKEN` partagé seul est un secret unique conservé dans le navigateur, non une authentification par utilisateur, et les identifiants d’utilisateur sont des déclarations. Il n’y a ni SSO, ni réinitialisation de mot de passe, ni contrôle d’accès par équipe et par canal. Voir [ce que le serveur de référence ne fait pas](/fr/server#ce-que-le-serveur-de-reference-ne-fait-pas).
- **La présence est indicative et non authentifiée**, comme toute awareness Yjs : un client peut
  revendiquer n’importe quel nom.
- Tout ce qu’envoie un pair est revalidé à la relecture et affiché comme du texte, jamais comme du
  balisage.

## Échelle

- **Les appels en maillage** connectent chaque participant à tous les autres. Cela ne nécessite
  aucun serveur média et convient à une poignée de personnes (le plafond par défaut est de 12),
  pas à un webinaire.
- **L’historique grossit.** Le serveur de référence conserve un instantané par salle et ne le
  compacte jamais.
- **Un seul processus serveur.** Le serveur de référence garde les salles en mémoire et ne passe
  pas à l’échelle horizontalement.

## Mode local et démos en direct

`mode: 'local'` utilise `BroadcastChannel`, donc seuls les onglets d’un même navigateur sur le même
site se voient. Rien n’est envoyé à un serveur, les fichiers ne sont partagés que par leur nom (il
n’y a nulle part où les envoyer), et le document partagé est conservé dans le `localStorage` de ce
navigateur. Les [démos en direct](/demos) fonctionnent ainsi parce que GitHub Pages ne peut pas
héberger de serveur.

## Appels entre réseaux

Un STUN public fonctionne sur les réseaux ouverts. Derrière un NAT symétrique ou des pare-feu
stricts, vous avez besoin de votre propre serveur TURN ; sans lui, certains appels n’aboutiront
pas. Les caméras et les microphones exigent `localhost` ou HTTPS.
