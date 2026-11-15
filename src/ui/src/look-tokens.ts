/**
 * The tokens of the opt-in "cards" look (`data-office-look="cards"`, see `look.ts`): shape,
 * elevation and density only, never a colour of the palette. They are part of the `--office-*`
 * contract and declared on `:root` like every other token, but the default look reads none of
 * them, so declaring them changes nothing until a scope opts into the look.
 */
export const LOOK_TOKENS = {
	/** A floating card: the ribbon, thumbnail rail, canvas, notes and side pane. */
	'--office-shadow-card': '0 1px 3px rgb(60 64 67 / 20%), 0 2px 6px rgb(60 64 67 / 8%)',
	/** A card above the cards: menus and popovers. */
	'--office-shadow-card-strong': '0 1px 3px rgb(60 64 67 / 30%), 0 4px 8px 3px rgb(60 64 67 / 15%)',
	/** Cards carry no hard border in light; dark themes give them a subtle one. */
	'--office-card-border': 'transparent',
	'--office-radius-surface': '14px',
	'--office-radius-toolbar': '12px',
	'--office-radius-control': '7px',
	/** The space between cards and around them. */
	'--office-workspace-gap': '12px',
	'--office-side-pane-width': '320px',
	'--office-pane-heading-height': '36px',
	'--office-pane-field-height': '32px',
} as const;

/** The dark values, declared by the theme's dark scopes with the rest of the palette. */
export const LOOK_DARK_TOKENS: Partial<Record<keyof typeof LOOK_TOKENS, string>> = {
	'--office-shadow-card': '0 1px 3px rgb(0 0 0 / 40%), 0 2px 6px rgb(0 0 0 / 20%)',
	'--office-shadow-card-strong': '0 1px 3px rgb(0 0 0 / 50%), 0 4px 8px 3px rgb(0 0 0 / 30%)',
	'--office-card-border': 'rgb(255 255 255 / 8%)',
};
