// Deutsche Fassung der Hinweise und Fehler, die der Kern auf Englisch formuliert.
import type { enCore } from '../en/core';

export const deCore: Record<keyof typeof enCore, string> = {
	'core.callFull': 'Der Anruf ist voll ({max} Teilnehmer)',
	'core.cameraUnavailable': 'Kamera nicht verfügbar',
	'core.cannotPin': 'Diese Nachricht kann nicht angeheftet werden',
	'core.cannotSaveFile': 'Diese Datei kann nicht im Kanal gespeichert werden',
	'core.channelGone': 'Der Kanal ist nicht mehr verfügbar',
	'core.channelGoneUploaded':
		'Der Kanal ist nicht mehr verfügbar; die hochgeladenen Dateien wurden nicht freigegeben',
	'core.chatNotCreated': 'Der Chat konnte nicht erstellt werden',
	'core.chooseFiles': 'Wählen Sie zwischen 1 und {max} Dateien',
	'core.configureStorage': 'Konfigurieren Sie den Dateispeicher, bevor Sie Dateien freigeben',
	'core.editNotApplied': 'Die Änderung wurde nicht übernommen; der Entwurf wurde beibehalten',
	'core.fileGone': 'Die Datei ist nicht mehr verfügbar',
	'core.fileNotUploaded': 'Die Datei konnte nicht hochgeladen werden',
	'core.filesNotShared': 'Die Dateien konnten nicht freigegeben werden',
	'core.invalidFileUrl': 'Der Server hat eine ungültige Datei-URL zurückgegeben',
	'core.linkFailed': 'Link fehlgeschlagen ({status})',
	'core.loadFailed': 'Der Inhalt konnte nicht geladen werden ({status})',
	'core.meetingInvalid':
		'Die Besprechung benötigt einen Titel und ein Ende nach ihrem Beginn (höchstens 24 Stunden)',
	'core.meetingUnavailable': 'Diese Besprechung ist nicht verfügbar',
	'core.openFailed': 'Die Datei konnte nicht geöffnet werden',
	'core.organizerCancel': 'Nur der Organisator kann eine Besprechung absagen',
	'core.organizerCancelOccurrence':
		'Nur der Organisator kann einen Termin einer Besprechung absagen',
	'core.organizerChange': 'Nur der Organisator kann eine Besprechung ändern',
	'core.organizerChangeOccurrence':
		'Nur der Organisator kann einen Termin einer Besprechung ändern',
	'core.personUnavailable': 'Diese Person ist für einen Chat nicht verfügbar',
	'core.pickPerson': 'Wählen Sie mindestens eine Person aus',
	'core.reattachDrafts':
		'Fügen Sie die fehlenden Entwurfsanlagen erneut an oder verwerfen Sie sie, bevor Sie senden',
	'core.sendCanceled':
		'Senden der Nachricht abgebrochen. Die Nachricht wurde unter Entwürfe gespeichert.',
	'core.sendFailed': 'Die Nachricht wurde nicht gesendet: {reason}. Unter Entwürfe gespeichert.',
	'core.streamingUnavailable': 'Streaming von Inhalten ist nicht verfügbar',
	'core.tooLarge': 'Der Inhalt ist für die Vorschau zu groß',
	'core.uploadFailed': 'Hochladen fehlgeschlagen ({status})',
	'core.workbookName': 'Geben Sie einen Namen für die Arbeitsmappe ein',
};
