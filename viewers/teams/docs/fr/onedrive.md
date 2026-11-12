# Liens OneDrive et documents Office

Pour une configuration pas à pas (enregistrement dans Microsoft Entra, connexion et un exemple
complet), commencez par le [guide de démarrage rapide OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

OpenTeams peut transporter un lien OneDrive vers un fichier DOCX, XLSX, PPTX ou VSDX. La suite Office
intégrée ouvre ce fichier avec l’éditeur de documents correspondant et la connexion partagée
`ooxml-core/embedding/onedrive`. OpenTeams est lui-même un espace de travail, et non un cinquième
éditeur de fichiers Office, ni Microsoft Teams.

La même distinction s’applique à `openteams-react-viewer`, `openteams-vue-viewer`,
`openteams-angular-viewer`, `openteams-svelte-viewer`, `openteams-solid-viewer`,
`openteams-vanilla-viewer` et à l’élément `ooxml-ui/teams` :

| Élément partagé                        | Où les modifications sont conservées                                                                      |
| -------------------------------------- | --------------------------------------------------------------------------------------------------------- |
| Lien OneDrive                          | Dans l’élément OneDrive d’origine, uniquement lorsque l’hôte autorisé enregistre via sa session OneDrive. |
| Lien vers un fichier de l’espace local | Dans l’espace de travail Office local ; ce n’est pas un lien cloud entre appareils.                       |
| Pièce jointe du serveur OpenTeams      | Dans le stockage de ce serveur ; une copie téléversée est indépendante de l’original OneDrive.            |

Le flux de partage des fichiers connectés de la suite demande un rôle explicite d’affichage ou de
modification et une portée de lien. Publier ou copier l’URL d’un élément existant n’accorde pas
à lui seul l’accès à une autre personne. Ne créez jamais un lien anonyme de remplacement lorsqu’un
locataire refuse un lien restreint. `openteams-server` n’acquiert pas d’identifiants Microsoft, ne
reflète pas les dossiers du drive et ne fusionne pas les modifications des fichiers Office.

Utilisez le [guide d’intégration OneDrive pour Office](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#sharing-and-openteams)
pour l’authentification, les références de paquets et le partage tenant compte des autorisations.
Son [guide de test](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) distingue
les transports locaux aux tests de la validation sur un vrai compte. Aucune interface ni aucun
service de simulation n’est fourni.

Les salles de discussion et de présence Yjs, ainsi que les salles de collaboration sur les documents,
sont distinctes de la coédition Microsoft Office. Envoyer la même URL OneDrive à deux personnes ne
synchronise pas automatiquement leurs éditeurs.
