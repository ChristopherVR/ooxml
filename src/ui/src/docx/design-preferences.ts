import {
	parseWordDesignDefaults,
	withWordDesignDefaults,
	type DocumentModel,
	type WordDesignDefaults,
	type WordDesignScheme,
} from 'ooxml-core/docx';
import type { EditorCore } from './editor-core';
import { loadDesignGallery } from './design-ribbon-controls';

/** Optional host-owned preferences for defaults and named custom schemes. */
export interface WordDesignPreferences {
	read(): unknown;
	write(value: WordDesignDefaults): void;
	schemes?: {
		read(): unknown;
		write(value: WordDesignScheme[]): void;
	};
}

export function wordDesignSeed(
	model: DocumentModel,
	preferences: WordDesignPreferences | undefined,
) {
	try {
		const settings = parseWordDesignDefaults(preferences?.read());
		return { model: withWordDesignDefaults(model, settings), effectName: settings?.effectName };
	} catch {
		return { model, effectName: undefined };
	}
}

export const WORD_DESIGN_DEFAULT = 'docx-design-default';

/** A saved effect can load after the editor mounts. Keep the fresh document's body and history. */
export function restoreWordDesignEffect(
	core: EditorCore,
	name: WordDesignDefaults['effectName'],
): void {
	const theme = core.model.theme;
	if (!name || !theme) return;
	const generation = core.loadGeneration;
	const pending = loadDesignGallery()
		.then(([{ applySavedDesignEffect }, { THEME_EFFECT_PRESETS }]) => {
			if (core.loadGeneration !== generation || core.model.theme !== theme || core.collab.active)
				return;
			const preset = THEME_EFFECT_PRESETS.find((preset) => preset.name === name);
			if (preset)
				applySavedDesignEffect(core, {
					...theme,
					formatScheme: structuredClone(preset.formatScheme),
				});
		})
		.catch(() => {})
		.finally(() => {
			if (core.pendingWordDesign === pending) core.pendingWordDesign = undefined;
		});
	core.pendingWordDesign = pending;
}
