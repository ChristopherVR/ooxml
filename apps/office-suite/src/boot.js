/**
 * The suite's entry point. On the root page it asks what to open before anything else loads:
 * the workspace shell stays hidden and the suite module (editors, storage, Teams) is fetched
 * only once the suite is picked. Product pages and links into a file load the suite at once.
 */
import { needsChooser, showChooser } from './start.js';
import { onNewVersion, registerWorker } from './updates.js';

// The site's Content Security Policy never allows `eval`. Zod (in the lazily loaded assistant and
// automation chunks) probes for it with `Function('')` when it builds its first object schema,
// which the browser reports as a violation; with its JIT off it skips the probe and validates the
// same way. Set before any chunk that loads Zod.
(globalThis.__zod_globalConfig ??= {}).jitless = true;

// Updates are checked from the first paint. While the chooser is up nothing can be lost, so a
// newer build simply reloads into itself; once the workspace is open, pwa.js offers the reload.
if (/^https?:$/.test(location.protocol)) {
	void registerWorker().catch(() => {});
	onNewVersion(() => {
		if ('start' in document.body.dataset) location.reload();
	});
}

const loadMain = () =>
	import('./main.js').catch((error) => {
		delete document.documentElement.dataset.boot;
		throw error;
	});

if (needsChooser()) {
	document.body.dataset.start = '';
	showChooser({
		required: true,
		onSuite: () => {
			document.documentElement.dataset.boot = 'pending';
			delete document.body.dataset.start;
			void loadMain();
		},
	});
	delete document.documentElement.dataset.boot;
} else void loadMain();
