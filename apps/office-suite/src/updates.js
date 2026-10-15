/**
 * Service-worker registration and the update signal, shared by boot.js (the start chooser) and
 * pwa.js (the workspace). Every shell asset is versioned and sw.js is fetched with no cache, so an
 * update is noticed on the next check: at load, when the tab becomes visible, hourly, or when the
 * user asks. The new worker takes over at once (see scripts/pwa/service-worker.js); what the
 * window does then is up to the caller: reload, or offer a reload.
 */
import { suiteBase } from './product.js';

/** The build this page was served from (scripts/build-pwa.mjs writes the meta). */
export const buildStamp =
	document.querySelector('meta[name="ooxml-build"]')?.getAttribute('content') ?? '';
const supported = 'serviceWorker' in navigator && /^https?:$/.test(location.protocol);
const CHECK_INTERVAL = 60 * 60 * 1000;
let registration;
const listeners = new Set();
let hadController = supported && !!navigator.serviceWorker.controller;

/** Registers the worker once and keeps checking for a newer build. Resolves to the registration. */
export function registerWorker() {
	if (!supported) return Promise.resolve(null);
	registration ??= navigator.serviceWorker
		.register(new URL('sw.js', suiteBase), { scope: suiteBase.pathname, updateViaCache: 'none' })
		.then((reg) => {
			document.addEventListener('visibilitychange', () => {
				if (document.visibilityState === 'visible') void reg.update().catch(() => {});
			});
			setInterval(() => void reg.update().catch(() => {}), CHECK_INTERVAL);
			return reg;
		});
	return registration;
}

/**
 * Runs the listener when a newer build has taken control of this window (not on the first
 * install, which only makes the current build available offline).
 */
export function onNewVersion(listener) {
	if (!supported) return () => {};
	listeners.add(listener);
	return () => listeners.delete(listener);
}
if (supported)
	navigator.serviceWorker.addEventListener('controllerchange', () => {
		if (!hadController) {
			hadController = true;
			return;
		}
		for (const listener of listeners) listener();
	});

/**
 * Asks the browser for a newer sw.js right now. Resolves to true when a new build was found (it
 * installs and the onNewVersion listeners fire shortly after), false when this build is current.
 */
export async function checkForUpdates() {
	const reg = await registerWorker();
	if (!reg) return false;
	await reg.update();
	return !!(reg.installing || reg.waiting);
}
