/**
 * The suite's entry point. On the root page it asks what to open before anything else loads:
 * the workspace shell stays hidden and the suite module (editors, storage, Teams) is fetched
 * only once the suite is picked. Product pages and links into a file load the suite at once.
 */
import { needsChooser, showChooser } from './start.js';

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
