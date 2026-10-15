import { onProfile } from './profile-state.js';
import { product, suiteBase } from './product.js';
import { $, choose, notify } from './ui.js';
import { buildStamp, checkForUpdates, onNewVersion, registerWorker } from './updates.js';

/** Reloads into the workspace (the root would ask what to open again), keeping the route. */
function reloadWorkspace() {
	if (product) {
		location.reload();
		return;
	}
	const url = new URL('?suite', suiteBase);
	url.hash = location.hash;
	location.replace(url);
}

/**
 * A newer build has taken over: reload when nothing would be lost, otherwise show a banner with
 * the choice. The banner stays until the reload; a reload is never forced over open work.
 * @param {() => boolean} idle whether a reload would lose nothing (no open documents)
 */
function offerReload(idle) {
	if (idle()) {
		reloadWorkspace();
		return;
	}
	if ($('update-banner')) return;
	const banner = document.createElement('div');
	banner.id = 'update-banner';
	banner.className = 'update-banner';
	banner.setAttribute('role', 'status');
	banner.innerHTML =
		'<span>A new version of OOXML Office is ready. Save your work, then reload to use it.</span><button class="primary" data-reload>Reload now</button><button data-later>Later</button>';
	banner.querySelector('[data-reload]').onclick = reloadWorkspace;
	banner.querySelector('[data-later]').onclick = () => {
		banner.remove();
		notify('A new version is ready. Reload whenever you like.');
	};
	document.querySelector('.suite-body').prepend(banner);
}

/** @param {{ idle: () => boolean }} options */
export function mountPwa({ idle }) {
	if (!('serviceWorker' in navigator) || !/^https?:$/.test(location.protocol)) return;
	let prompt;
	const install = document.createElement('button');
	install.textContent = 'Install this app';
	install.className = 'install-app';
	const check = document.createElement('button');
	check.textContent = 'Check for updates';
	check.className = 'check-updates';
	const version = document.createElement('p');
	version.className = 'account-version';
	version.textContent = buildStamp ? `Version ${buildStamp}` : 'Development build';
	onProfile(() => document.querySelector('#account-menu').append(install, check, version));
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
	check.onclick = async () => {
		document.querySelector('#account-menu').hidePopover();
		check.disabled = true;
		notify('Checking for updates...');
		try {
			const found = await checkForUpdates();
			notify(found ? 'A new version is downloading.' : 'You have the latest version.');
		} catch {
			notify('Could not check for updates. Try again when you are online.');
		} finally {
			check.disabled = false;
		}
	};
	window.addEventListener('appinstalled', () => {
		install.hidden = true;
	});
	onNewVersion(() => offerReload(idle));
	const register = () =>
		registerWorker().catch(() =>
			notify('Offline setup is unavailable. The app remains usable online.'),
		);
	if (document.readyState === 'complete') void register();
	else window.addEventListener('load', () => void register(), { once: true });
}
