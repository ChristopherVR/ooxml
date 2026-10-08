// Send a returning visitor straight to the app they picked on the start chooser
// (apps/office-suite/src/start.js). Runs before paint so the suite never flashes first.
try {
	const app = localStorage.getItem('ooxml-start-app') ?? '';
	if (
		document.querySelector('meta[name="ooxml-root"]')?.content === './' &&
		/^https?:$/.test(location.protocol) &&
		/^(word|excel|powerpoint|visio|teams)$/.test(app) &&
		(!location.hash || location.hash === '#/') &&
		!new URLSearchParams(location.search).has('suite') &&
		!matchMedia('(display-mode: standalone)').matches
	)
		location.replace(`apps/${app}/`);
} catch {}
