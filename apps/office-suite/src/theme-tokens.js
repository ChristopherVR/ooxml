/** Each accent: its name, the light-mode value and the dark-mode value. */
export const accents = {
	blue: ['Office blue', '#0f6cbd', '#479ef5'],
	purple: ['Violet', '#5b5fc7', '#a3a5f0'],
	teal: ['Teal', '#03787c', '#4cc2c7'],
	green: ['Forest', '#107c41', '#54b054'],
	orange: ['Amber', '#a84b00', '#f7924a'],
	rose: ['Rose', '#b4126e', '#f27ac2'],
	cranberry: ['Cranberry', '#a4262c', '#f1707b'],
	graphite: ['Graphite', '#525252', '#a8a8a8'],
};
/**
 * The accent as text on the canvas and on selected fills: the same in light mode, lifted in dark
 * mode so it reads at 4.5:1 (scripts/suite-theme-contrast.test.mjs checks every pair).
 */
export const accentInks = {
	blue: ['#0f6cbd', '#7ab8ff'],
	purple: ['#5b5fc7', '#bdbef8'],
	teal: ['#03787c', '#7ad8dc'],
	green: ['#107c41', '#7ccf8a'],
	orange: ['#a84b00', '#ffae73'],
	rose: ['#b4126e', '#f7a3d6'],
	cranberry: ['#a4262c', '#f7a1a8'],
	graphite: ['#525252', '#c4c4c4'],
};

/** The resolved theme for a mode ('light' or 'dark') and an accent key of `accents`. */
export function resolveTheme(mode, accentKey) {
	const dark = mode === 'dark' ? 1 : 0;
	return { mode, accent: accents[accentKey][1 + dark], accentInk: accentInks[accentKey][dark] };
}

/** One suite palette, adapted to each editor's public theme API. */
export function editorColors({ mode, accent, accentInk = accent }) {
	const dark = mode === 'dark';
	const surface = dark ? '#202228' : '#ffffff';
	const text = dark ? '#f0f0f0' : '#242424';
	return {
		background: surface,
		foreground: text,
		card: surface,
		cardForeground: text,
		popover: surface,
		popoverForeground: text,
		primary: accent,
		primaryForeground: dark ? '#141414' : '#ffffff',
		primaryInk: accentInk,
		secondary: dark ? '#292d34' : '#f0f2f5',
		secondaryForeground: text,
		muted: dark ? '#181a1e' : '#f7f8fa',
		mutedForeground: dark ? '#bbbbbb' : '#616161',
		accent: dark ? '#303c49' : '#e7eef8',
		accentForeground: text,
		border: dark ? '#34383f' : '#e3e6eb',
		// Form control boundaries: 3:1 against every surface above. Dividers keep `border`.
		input: '#8a8886',
		ring: accent,
		destructive: dark ? '#ff8585' : '#b42318',
		destructiveForeground: '#ffffff',
	};
}
/** The `--office-*` values the suite sets on its root and on each editor. */
export function officeThemeValues(theme) {
	const c = editorColors(theme);
	// Hover tokens are left to their shared `color-mix()` defaults, which follow the accent.
	return {
		foreground: c.foreground,
		'muted-foreground': c.mutedForeground,
		background: c.background,
		surface: c.secondary,
		border: c.border,
		'control-border': c.input,
		accent: c.primary,
		'accent-foreground': c.primaryForeground,
		'accent-ink': c.primaryInk,
		ring: c.ring,
		selected: c.accent,
		popover: c.popover,
		'field-background': c.background,
		'teams-brand': c.primary,
		'teams-brand-foreground': c.primaryForeground,
		'teams-brand-ink': c.primaryInk,
		'teams-brand-subtle': c.accent,
		'teams-rail': c.muted,
		'teams-panel': c.secondary,
		'teams-bubble': c.background,
		'teams-bubble-own': c.accent,
		'teams-divider': c.border,
		'teams-text-subtle': c.mutedForeground,
	};
}
export function applyOfficeTheme(element, theme) {
	for (const [name, value] of Object.entries(officeThemeValues(theme)))
		element.style.setProperty(`--office-${name}`, value);
	element.setAttribute('data-office-theme', theme.mode);
	element.style.colorScheme = theme.mode;
}

/** Floating PowerPoint menus/dialogs are attached to body, outside their viewer. */
export function applyPowerPointTheme(element, theme) {
	for (const [name, value] of Object.entries(editorColors(theme)))
		element.style.setProperty(
			`--pptx-${name.replace(/[A-Z]/g, (c) => '-' + c.toLowerCase())}`,
			value,
		);
	element.style.setProperty('--pptx-radius', '4px');
	element.style.setProperty('--pptx-inspector-active', theme.accent);
	element.style.setProperty('--pptx-inspector-border', editorColors(theme).border);
}
