// Apply the saved mode before styles paint. Full theme controls live in the suite host.
try {
	const preference = JSON.parse(localStorage.getItem('ooxml-suite-theme') ?? 'null');
	const mode = preference?.mode;
	const dark =
		mode === 'dark' || (mode !== 'light' && matchMedia('(prefers-color-scheme: dark)').matches);
	document.documentElement.dataset.theme = dark ? 'dark' : 'light';
	document.documentElement.style.colorScheme = dark ? 'dark' : 'light';
} catch {}
