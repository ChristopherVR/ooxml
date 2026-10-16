// Versión en español de los avisos y errores que el núcleo redacta en inglés.
import type { enCore } from '../en/core';

export const esCore: Record<keyof typeof enCore, string> = {
	'core.callFull': 'La llamada está llena ({max} participantes)',
	'core.cameraUnavailable': 'Cámara no disponible',
	'core.cannotPin': 'Este mensaje no se puede anclar',
	'core.cannotSaveFile': 'Este archivo no se puede guardar en el canal',
	'core.channelGone': 'El canal ya no está disponible',
	'core.channelGoneUploaded':
		'El canal ya no está disponible; los archivos cargados no se compartieron',
	'core.chatNotCreated': 'No se pudo crear el chat',
	'core.chooseFiles': 'Elige entre 1 y {max} archivos',
	'core.configureStorage': 'Configura el almacenamiento de archivos antes de compartir archivos',
	'core.editNotApplied': 'No se aplicó la edición; se conservó el borrador',
	'core.fileGone': 'El archivo ya no está disponible',
	'core.fileNotUploaded': 'No se pudo cargar el archivo',
	'core.filesNotShared': 'No se pudieron compartir los archivos',
	'core.invalidFileUrl': 'El servidor devolvió una dirección URL de archivo no válida',
	'core.linkFailed': 'Error en el vínculo ({status})',
	'core.loadFailed': 'No se pudo cargar el contenido ({status})',
	'core.meetingInvalid':
		'La reunión necesita un título y un fin posterior a su inicio (24 horas como máximo)',
	'core.meetingUnavailable': 'Esta reunión no está disponible',
	'core.openFailed': 'No se pudo abrir el archivo',
	'core.organizerCancel': 'Solo el organizador puede cancelar una reunión',
	'core.organizerCancelOccurrence':
		'Solo el organizador puede cancelar una repetición de una reunión',
	'core.organizerChange': 'Solo el organizador puede cambiar una reunión',
	'core.organizerChangeOccurrence':
		'Solo el organizador puede cambiar una repetición de una reunión',
	'core.personUnavailable': 'Esa persona no está disponible para chatear',
	'core.pickPerson': 'Elige al menos una persona',
	'core.reattachDrafts':
		'Vuelve a adjuntar o descarta los archivos que faltan en el borrador antes de enviar',
	'core.sendCanceled': 'Se canceló el envío del mensaje. El mensaje se guardó en Borradores.',
	'core.sendFailed': 'No se envió el mensaje: {reason}. Guardado en Borradores.',
	'core.streamingUnavailable': 'La transmisión de contenido no está disponible',
	'core.tooLarge': 'El contenido es demasiado grande para la vista previa',
	'core.uploadFailed': 'Error al cargar ({status})',
	'core.workbookName': 'Escribe un nombre para el libro',
};
