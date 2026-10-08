/** One suite palette, adapted to each editor's public theme API. */
export function editorColors({ mode, accent }) {
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
		secondary: dark ? '#292d34' : '#f0f2f5',
		secondaryForeground: text,
		muted: dark ? '#181a1e' : '#f7f8fa',
		mutedForeground: dark ? '#bbbbbb' : '#616161',
		accent: dark ? '#303c49' : '#e7eef8',
		accentForeground: text,
		border: dark ? '#34383f' : '#e3e6eb',
		input: dark ? '#4d535d' : '#b8bec8',
		ring: accent,
		destructive: dark ? '#ff8585' : '#b42318',
		destructiveForeground: '#ffffff',
	};
}
export function applyOfficeTheme(element, theme) {
	const c = editorColors(theme);
	const values = {
		foreground: c.foreground,
		'muted-foreground': c.mutedForeground,
		background: c.background,
		surface: c.secondary,
		border: c.border,
		accent: c.primary,
		'accent-hover': c.primary,
		'accent-foreground': c.primaryForeground,
		ring: c.ring,
		selected: c.accent,
		popover: c.popover,
		'field-background': c.background,
		'teams-brand': c.primary,
		'teams-brand-hover': c.primary,
		'teams-brand-subtle': c.accent,
		'teams-rail': c.muted,
		'teams-panel': c.secondary,
		'teams-bubble': c.background,
		'teams-bubble-own': c.accent,
		'teams-divider': c.border,
		'teams-text-subtle': c.mutedForeground,
	};
	for (const [name, value] of Object.entries(values))
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
