// Version française des avis et erreurs que le cœur formule en anglais.
import type { enCore } from '../en/core';

export const frCore: Record<keyof typeof enCore, string> = {
	'core.callFull': "L'appel est complet ({max} participants)",
	'core.cameraUnavailable': 'Caméra indisponible',
	'core.cannotPin': 'Ce message ne peut pas être épinglé',
	'core.cannotSaveFile': 'Ce fichier ne peut pas être enregistré dans le canal',
	'core.channelGone': "Le canal n'est plus disponible",
	'core.channelGoneUploaded':
		"Le canal n'est plus disponible ; les fichiers chargés n'ont pas été partagés",
	'core.chatNotCreated': 'Impossible de créer la conversation',
	'core.chooseFiles': 'Choisissez entre 1 et {max} fichiers',
	'core.configureStorage': 'Configurez le stockage de fichiers avant de partager des fichiers',
	'core.editNotApplied': "La modification n'a pas été appliquée ; le brouillon a été conservé",
	'core.fileGone': "Le fichier n'est plus disponible",
	'core.fileNotUploaded': 'Impossible de charger le fichier',
	'core.filesNotShared': 'Impossible de partager les fichiers',
	'core.invalidFileUrl': 'Le serveur a renvoyé une URL de fichier non valide',
	'core.linkFailed': 'Échec du lien ({status})',
	'core.loadFailed': 'Impossible de charger le contenu ({status})',
	'core.meetingInvalid':
		'La réunion nécessite un titre et une fin postérieure à son début (24 heures au maximum)',
	'core.meetingUnavailable': "Cette réunion n'est pas disponible",
	'core.openFailed': "Impossible d'ouvrir le fichier",
	'core.organizerCancel': "Seul l'organisateur peut annuler une réunion",
	'core.organizerCancelOccurrence': "Seul l'organisateur peut annuler une occurrence d'une réunion",
	'core.organizerChange': "Seul l'organisateur peut modifier une réunion",
	'core.organizerChangeOccurrence':
		"Seul l'organisateur peut modifier une occurrence d'une réunion",
	'core.personUnavailable': "Cette personne n'est pas disponible pour discuter",
	'core.pickPerson': 'Choisissez au moins une personne',
	'core.reattachDrafts':
		"Joignez à nouveau ou abandonnez les pièces jointes manquantes du brouillon avant l'envoi",
	'core.sendCanceled': 'Envoi du message annulé. Le message a été enregistré dans Brouillons.',
	'core.sendFailed': "Le message n'a pas été envoyé : {reason}. Enregistré dans Brouillons.",
	'core.streamingUnavailable': 'La diffusion du contenu est indisponible',
	'core.tooLarge': 'Le contenu est trop volumineux pour être affiché en aperçu',
	'core.uploadFailed': 'Échec du chargement ({status})',
	'core.workbookName': 'Saisissez un nom de classeur',
};
