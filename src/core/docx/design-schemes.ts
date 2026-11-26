import { DESIGN_COLOR_SLOTS, type DesignEdit } from './document-design';
import { boundedDesignText, isDesignRecord } from './design-defaults';

export type WordDesignScheme = Extract<DesignEdit, { kind: 'colors' | 'fonts' }>;
export const MAX_WORD_DESIGN_SCHEMES = 24;

function parseScheme(value: unknown): WordDesignScheme | undefined {
	if (!isDesignRecord(value)) return;
	const name = boundedDesignText(value.name);
	if (!name) return;
	if (value.kind === 'fonts') {
		const majorFont = boundedDesignText(value.majorFont);
		const minorFont = boundedDesignText(value.minorFont);
		if (majorFont && minorFont) return { kind: 'fonts', name, majorFont, minorFont };
	}
	if (value.kind !== 'colors' || !isDesignRecord(value.colors)) return;
	const colors: Extract<WordDesignScheme, { kind: 'colors' }>['colors'] = {};
	for (const slot of DESIGN_COLOR_SLOTS) {
		const color = value.colors[slot];
		if (typeof color !== 'string' || !/^[0-9A-F]{6}$/i.test(color)) return;
		colors[slot] = color.toUpperCase();
	}
	return { kind: 'colors', name, colors };
}

export function sameWordDesignScheme(a: WordDesignScheme, b: WordDesignScheme): boolean {
	return a.kind === b.kind && a.name.toLowerCase() === b.name.toLowerCase();
}

/** Only named colour palettes and Latin font pairs are admitted from the host preference store. */
export function parseWordDesignSchemes(value: unknown): WordDesignScheme[] {
	if (!Array.isArray(value)) return [];
	const schemes: WordDesignScheme[] = [];
	for (const candidate of value.slice(0, MAX_WORD_DESIGN_SCHEMES)) {
		const scheme = parseScheme(candidate);
		if (scheme && !schemes.some((item) => sameWordDesignScheme(item, scheme))) schemes.push(scheme);
	}
	return schemes;
}
