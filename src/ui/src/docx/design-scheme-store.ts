import {
	MAX_WORD_DESIGN_SCHEMES,
	parseWordDesignSchemes,
	sameWordDesignScheme,
	type WordDesignScheme,
} from 'ooxml-core/docx';
import type { EditorCore } from './editor-core';
import { translate } from './localization';

export function readDesignSchemes(core: EditorCore): WordDesignScheme[] {
	try {
		return parseWordDesignSchemes(core.wordDesignPreferences?.schemes?.read());
	} catch {
		return [];
	}
}

export function changeDesignScheme(
	core: EditorCore,
	previous: WordDesignScheme | undefined,
	next: WordDesignScheme | undefined,
): boolean {
	const store = core.wordDesignPreferences?.schemes;
	if (!store) return false;
	try {
		const schemes = parseWordDesignSchemes(store.read());
		const index = previous ? schemes.findIndex((item) => sameWordDesignScheme(item, previous)) : -1;
		if (previous && index < 0) throw new Error();
		if (next) {
			const [parsed] = parseWordDesignSchemes([next]);
			if (
				!parsed ||
				schemes.some((item, at) => at !== index && sameWordDesignScheme(item, parsed)) ||
				(index < 0 && schemes.length >= MAX_WORD_DESIGN_SCHEMES)
			)
				throw new Error();
			if (index < 0) schemes.push(parsed);
			else schemes[index] = parsed;
		} else if (index >= 0) schemes.splice(index, 1);
		store.write(schemes);
		return true;
	} catch {
		core.host.reportError(translate(core.locale, 'design.schemeError'));
		return false;
	}
}
