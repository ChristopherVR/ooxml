import {
	withWordDesignDefaults,
	type DocumentModel,
	type WordDesignDefaults,
	type WordDesignScheme,
} from 'ooxml-core/docx';

/** Optional host-owned preferences for defaults and named custom schemes. */
export interface WordDesignPreferences {
	read(): unknown;
	write(value: WordDesignDefaults): void;
	schemes?: {
		read(): unknown;
		write(value: WordDesignScheme[]): void;
	};
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
