import { onProfile } from './profile-state.js';
import { suiteBase } from './product.js';
import { choose, notify } from './ui.js';
/** Tell the user a newer version is ready; it loads on the next reload. */
function offerReload() {
	notify('A new version of OOXML Office is ready.');
	const reload = document.createElement('button');
	reload.className = 'update-app';
	reload.textContent = 'Reload to update';
	reload.onclick = () => location.reload();
	document.querySelector('#status')?.after(reload);
}
export function mountPwa() {
	if (!('serviceWorker' in navigator) || !/^https?:$/.test(location.protocol)) return;
	let prompt;
	const install = document.createElement('button');
	install.textContent = 'Install this app';
	install.className = 'install-app';
	onProfile(() => document.querySelector('#account-menu').append(install));
	window.addEventListener('beforeinstallprompt', (event) => {
		event.preventDefault();
		prompt = event;
	});
	install.onclick = async () => {
		document.querySelector('#account-menu').hidePopover();
		if (prompt) {
			await prompt.prompt();
			prompt = null;
			return;
		}
		const message = document.createElement('p');
		message.textContent =
			'Use Install app in your browser menu. On iPhone or iPad, open this page in Safari, choose Share, then Add to Home Screen. Each app can be installed from its own page in the app launcher.';
		choose('Install app', message);
	};
	window.addEventListener('appinstalled', () => {
		install.hidden = true;
	});
	// A new worker takes over as soon as it is installed; this window keeps the code it loaded.
	if (navigator.serviceWorker.controller)
		navigator.serviceWorker.addEventListener('controllerchange', offerReload, { once: true });
	const register = () =>
		navigator.serviceWorker
			// sw.js always comes from the network, so a deploy is noticed on the next check.
			.register(new URL('sw.js', suiteBase), { scope: suiteBase.pathname, updateViaCache: 'none' })
			.then((registration) =>
				document.addEventListener('visibilitychange', () => {
					if (document.visibilityState === 'visible') void registration.update().catch(() => {});
				}),
			)
			.catch(() => notify('Offline setup is unavailable. The app remains usable online.'));
	if (document.readyState === 'complete') void register();
	else window.addEventListener('load', () => void register(), { once: true });
}
