import { onProfile } from './profile-state.js';
import { suiteBase } from './product.js';
import { choose, notify } from './ui.js';
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
	const register = () =>
		navigator.serviceWorker
			.register(new URL('sw.js', suiteBase), { scope: suiteBase.pathname })
			.catch(() => notify('Offline setup is unavailable. The app remains usable online.'));
	if (document.readyState === 'complete') void register();
	else window.addEventListener('load', () => void register(), { once: true });
}
