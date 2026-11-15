/**
 * The Office chrome's looks. `data-office-theme` picks the palette (light, dark, auto); the look,
 * `data-office-look` on the same scope, picks shape, elevation and density and never a colour.
 * The default look is today's flat chrome. `cards` floats the chrome regions (ribbon, thumbnail
 * rail, canvas, notes, side pane) as rounded, shadowed cards on a workspace background, with
 * rounder controls, pill switchers and a roomier Format pane.
 *
 * Elements in shadow roots cannot see the attribute, so the look reaches them through private
 * `--_office-look-*` properties declared on the look's scope. Each element reads one with its
 * default-look value as the fallback, so nothing changes outside the look, and the look never
 * redeclares a public token a host may have overridden. A product's light-DOM chrome selects the
 * attribute directly (for PowerPoint, `pptx/render/editor-chrome/cards-css.ts`).
 */
export const OFFICE_LOOK_ATTRIBUTE = 'data-office-look';

export const OFFICE_LOOKS = ['default', 'cards'] as const;
export type OfficeLook = (typeof OFFICE_LOOKS)[number];

/** The scopes the cards look applies to: its own, and every theme scope inside it. */
const CARDS_SCOPE =
	':where([data-office-look="cards"], [data-office-look="cards"] [data-office-theme])';

/**
 * The private properties the shared elements read under the cards look. Declared again on each
 * theme scope inside the look, so a dark scope nested in a light one gets its own shadows.
 */
export const CARDS_LOOK_PROPERTIES: Readonly<Record<string, string>> = {
	'--_office-look-menu-radius': 'var(--office-radius-toolbar, 12px)',
	'--_office-look-menu-shadow':
		'var(--office-shadow-card-strong, 0 1px 3px rgb(60 64 67 / 30%), 0 4px 8px 3px rgb(60 64 67 / 15%))',
	'--_office-look-pill': 'var(--office-radius-full, 999px)',
	// `currentColor` resolves where the property is used: a neutral tint of the text there.
	'--_office-look-pill-track': 'color-mix(in srgb, currentColor 9%, transparent)',
	'--_office-look-control-radius': 'var(--office-radius-control, 7px)',
	'--_office-look-heading-size': 'var(--office-font-size, 13px)',
	'--_office-look-heading-height': 'var(--office-pane-heading-height, 36px)',
};

export const CARDS_LOOK_CSS = `${CARDS_SCOPE} {
${Object.entries(CARDS_LOOK_PROPERTIES)
	.map(([name, value]) => `\t${name}: ${value};`)
	.join('\n')}
}
`;

/** The look a value names, `default` for anything else (a missing or unknown attribute). */
export function normalizeOfficeLook(value: unknown): OfficeLook {
	return value === 'cards' ? 'cards' : 'default';
}

/**
 * Set a scope's look. The default look removes the attribute, so the scope inherits whatever an
 * outer scope chose. Only the attribute changes: mounted editors keep their state.
 */
export function applyOfficeLook(element: Element, look: OfficeLook | undefined): void {
	if (normalizeOfficeLook(look) === 'default') element.removeAttribute(OFFICE_LOOK_ATTRIBUTE);
	else element.setAttribute(OFFICE_LOOK_ATTRIBUTE, 'cards');
}
