/** Stable ids for cards in the no-selection property inspector. */
export const INSPECTOR_CARD_IDS = [
	'presentation',
	'theme',
	'themeOverride',
	'slideSize',
	'slideTransition',
	'notesHandout',
	'document',
	'tags',
	'slide',
	'slideBackground',
	'templateBackground',
] as const;

export type InspectorCardId = (typeof INSPECTOR_CARD_IDS)[number];

/** Preserve each binding's desktop/mobile default when the host omits the option. */
export function initialInspectorOpen(option: boolean | undefined, defaultOpen = true): boolean {
	return option ?? defaultOpen;
}
