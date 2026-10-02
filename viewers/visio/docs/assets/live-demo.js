const button = document.getElementById('load-demo');
const frame = document.getElementById('live-viewer');
const status = document.getElementById('demo-status');
let timer;
let started = false;
let ready = false;
button?.addEventListener('click', () => {
	if (started || !frame || !status) return;
	started = true;
	button.hidden = true;
	frame.hidden = false;
	frame.setAttribute('aria-busy', 'true');
	status.textContent = 'Loading the live viewer…';
	frame.src = frame.dataset.src;
	timer = setTimeout(() => {
		if (ready) return;
		// Leave the frame intact so a slow startup can still finish without losing state.
		frame.setAttribute('aria-busy', 'false');
		status.textContent =
			'The viewer is taking longer than expected. You can use the full playground link above.';
	}, 15000);
});
window.addEventListener('message', (event) => {
	if (
		!started ||
		ready ||
		event.origin !== location.origin ||
		event.source !== frame?.contentWindow
	)
		return;
	if (event.data?.type !== 'visio-viewer-ready') return;
	ready = true;
	clearTimeout(timer);
	try {
		if (frame.contentDocument)
			frame.contentDocument.documentElement.dataset.theme = document.documentElement.dataset.theme;
	} catch {
		// The configured viewer is same-origin; never touch another document.
	}
	frame.setAttribute('aria-busy', 'false');
	status.textContent = 'Viewer ready. Open a local .vsdx or explore the sample.';
});
window.addEventListener('pagehide', () => clearTimeout(timer));
