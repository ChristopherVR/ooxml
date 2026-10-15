/**
 * The suite's entry point. On the root page it asks what to open before anything else loads:
 * the workspace shell stays hidden and the suite module (editors, storage, Teams) is fetched
 * only once the suite is picked. Product pages and links into a file load the suite at once.
 */
import { needsChooser, showChooser } from './start.js';
import { onNewVersion, registerWorker } from './updates.js';

// Updates are checked from the first paint. While the chooser is up nothing can be lost, so a
// newer build simply reloads into itself; once the workspace is open, pwa.js offers the reload.
if (/^https?:$/.test(location.protocol)) {
	void registerWorker().catch(() => {});
	onNewVersion(() => {
		if ('start' in document.body.dataset) location.reload();
	});
}

if (needsChooser()) {
	document.body.dataset.start = '';
	showChooser({
		required: true,
		onSuite: () => {
			delete document.body.dataset.start;
			void import('./main.js');
		},
	});
} else void import('./main.js');
