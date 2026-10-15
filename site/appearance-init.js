// Apply the saved mode before styles paint. Full theme controls live in the suite host.
// Without a saved preference the workspace opens dark; "system" follows the device.
try {
	const preference = JSON.parse(localStorage.getItem('ooxml-suite-theme') ?? 'null');
	const mode = preference?.mode;
	const dark =
		mode === 'dark' ||
		(mode === 'system' && matchMedia('(prefers-color-scheme: dark)').matches) ||
		(mode !== 'light' && mode !== 'system');
	document.documentElement.dataset.theme = dark ? 'dark' : 'light';
	document.documentElement.style.colorScheme = dark ? 'dark' : 'light';
} catch {}
