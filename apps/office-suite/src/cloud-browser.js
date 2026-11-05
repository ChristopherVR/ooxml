import { $, escape, notify, task } from './ui.js';

/**
 * The "OneDrive" workspace view: the shared `<office-ui-cloud-files>` browser (Recent, My files with
 * folders, search, open from a link) for all four Office kinds. Choosing a file hands the download
 * to the suite's cloud flow, so it opens in its editor with the same identity, Save, conflict
 * recovery and Share as a file opened from a link. A file already in the library reuses its document.
 */
export function mountCloudBrowser({ cloud, options, open, refresh }) {
	const host = $('cloud-view');
	let element;
	const render = async () => {
		if (element) return;
		host.innerHTML = `<div class="cloud-view-head"><h1>${escape(options.label)}</h1><p>Browse your ${escape(options.label)} files and open them in Word, Excel, PowerPoint or Visio. Save updates the same file. Live Microsoft coauthoring is not enabled.</p></div>`;
		// The element and its Lit runtime load on first use, not with the shell.
		const { defineCloudFiles } = await import('ooxml-ui/controls');
		defineCloudFiles();
		element = document.createElement('office-ui-cloud-files');
		element.id = 'cloud-files';
		element.providers = [options.fileProvider];
		element.locale = document.documentElement.lang || undefined;
		element.addEventListener('office-cloud-file-open', (event) =>
			task(async () => {
				const source = options.sourceOf(event.detail.file);
				if (!source) throw new Error(`${options.label} could not open this file.`);
				const doc = await cloud.adopt(source);
				await open(doc.id, undefined, { focus: true });
				await refresh();
			}),
		);
		host.append(element);
	};
	return {
		/** Shows the browser, mounting it on first use. */
		async show() {
			try {
				await render();
			} catch (error) {
				notify(error);
			}
		},
	};
}
