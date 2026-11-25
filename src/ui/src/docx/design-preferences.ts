import {
	withWordDesignDefaults,
	type DocumentModel,
	type WordDesignDefaults,
} from 'ooxml-core/docx';

/** Optional host-owned preferences. Omit this property to disable reading and writing defaults. */
export interface WordDesignPreferences {
	read(): unknown;
	write(value: WordDesignDefaults): void;
}

export function applyWordDesignPreferences(
	model: DocumentModel,
	preferences: WordDesignPreferences | undefined,
): DocumentModel {
	try {
		return withWordDesignDefaults(model, preferences?.read());
	} catch {
		return model;
	}
}
