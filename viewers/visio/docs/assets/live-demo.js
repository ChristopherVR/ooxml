const button = document.getElementById('load-demo');
const frame = document.getElementById('live-viewer');
const status = document.getElementById('demo-status');
let timer;
button?.addEventListener('click', () => {
	button.hidden = true;
	frame.hidden = false;
	status.textContent = 'Loading the live viewer…';
	frame.src = frame.dataset.src;
	timer = setTimeout(() => {
		status.textContent = 'The viewer did not become ready. Try the full playground link above.';
	}, 15000);
});
window.addEventListener('message', (event) => {
	if (event.origin !== location.origin || event.source !== frame?.contentWindow) return;
	if (event.data?.type !== 'visio-viewer-ready') return;
	clearTimeout(timer);
	status.textContent = 'Viewer ready. Open a local .vsdx or explore the sample.';
});
