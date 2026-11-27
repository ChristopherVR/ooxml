// Apply the saved mode before styles paint. Full theme controls live in the suite host.
// Without a saved preference the workspace opens dark; "system" follows the device. A page whose
// root sets data-appearance-default="system" (the release posts) follows the device until the
// reader picks a mode.
try {
	const preference = JSON.parse(localStorage.getItem('ooxml-suite-theme') ?? 'null');
	const mode = preference?.mode;
	const followsSystem =
		mode === 'system' ||
		(mode !== 'light' &&
			mode !== 'dark' &&
			document.documentElement.dataset.appearanceDefault === 'system');
	const dark =
		mode === 'dark' ||
		(followsSystem && matchMedia('(prefers-color-scheme: dark)').matches) ||
		(!followsSystem && mode !== 'light');
	document.documentElement.dataset.theme = dark ? 'dark' : 'light';
	document.documentElement.style.colorScheme = dark ? 'dark' : 'light';
	if (document.querySelector('meta[name="ooxml-root"]'))
		document.documentElement.dataset.boot = 'pending';
} catch {}
