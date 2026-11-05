import { SuiteOneDrive } from './cloud-files.js';
import { mountCloudBrowser } from './cloud-browser.js';
import { resolveCloudOptions } from './cloud-options.js';
import { $, choose, download, escape, notify, task } from './ui.js';

/** Shared shell controls for Word, Excel, PowerPoint and Visio, including standalone apps. */
export function mountOneDrive({ store, sessions, open, refresh, showView }) {
	const options = resolveCloudOptions();
	const label = options.label;
	const message = (text) => text.replaceAll('OneDrive', () => label);
	const html = (text) => text.replaceAll('OneDrive', () => escape(label));
	const cloud = new SuiteOneDrive({
		...options,
		store,
		changed: () => {
			void refresh();
		},
		adopted: (doc) => sessions.reload(doc),
		dirty: (id) => sessions.isDirty(id),
		conflict:
			options.resolveConflict ??
			(({ doc, conflict }) =>
				new Promise((resolve) => {
					const content = document.createElement('div');
					content.id = 'onedrive-conflict';
					content.innerHTML = `<p>Some edits in <strong>${escape(doc.name)}</strong> could not be combined automatically. Independent changes are merged during Save; the remaining changes overlap or cannot be matched safely. Your edits are saved on this device.</p>
				<p>The OneDrive version is ${escape(conflict.remote.file.name)}. Saving a copy keeps the original unchanged. Opening the latest first keeps your edits as a separate file in your library.</p>
				<div class="cloud-recovery-actions"><button data-cloud-recovery="copy" class="primary">Save my edits as a OneDrive copy</button><button data-cloud-recovery="latest">Keep local copy and open latest</button><button data-cloud-recovery="download">Download my edits</button><button data-cloud-recovery="download-latest">Download latest OneDrive version</button><button data-cloud-recovery="cancel">Keep editing</button></div>`;
					content.innerHTML = html(content.innerHTML);
					content.querySelector('[data-cloud-recovery="copy"]').disabled =
						!options.client?.saveCopy;
					if (conflict.remote.file.kind !== doc.kind) {
						content.querySelector('[data-cloud-recovery="latest"]').disabled = true;
						const note = document.createElement('p');
						note.textContent =
							'The cloud file changed format. Save your edits as a separate cloud copy before reopening that link.';
						content.append(note);
					}
					const dialog = choose('Resolve overlapping edits', content);
					let done = false;
					const finish = (action) => {
						if (done) return;
						done = true;
						dialog.removeEventListener('close', closed);
						dialog.close();
						resolve(action);
					};
					const closed = () => finish('cancel');
					dialog.addEventListener('close', closed);
					content.addEventListener('click', (event) => {
						const action = event.target.closest('[data-cloud-recovery]')?.dataset.cloudRecovery;
						if (action === 'download') download(doc);
						else if (action === 'download-latest')
							download({
								name: conflict.remote.file.name.replace(/(\.[^.]+)$/u, ' (OneDrive version)$1'),
								bytes: conflict.remote.bytes,
							});
						else if (action) finish(action);
					});
				})),
	});
	cloud.showOpen = () => {
		const content = document.createElement('form');
		content.id = 'onedrive-open-form';
		content.innerHTML = cloud.configured()
			? '<p>Open a Word, Excel, PowerPoint or Visio file using your Microsoft access. Save updates the connected file. Live Microsoft coauthoring is not enabled.</p><label>OneDrive sharing link<input id="onedrive-link" name="link" type="url" required autocomplete="off" spellcheck="false" /></label><button id="onedrive-open-submit" class="primary">Open from OneDrive</button>'
			: '<p>Microsoft sign-in has not been connected by your host or administrator. Ask them to enable OneDrive access. You can continue opening files from this device.</p><p><a href="https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#integrated-office-suite" target="_blank" rel="noopener noreferrer">Integration documentation</a></p>';
		if (options.id !== 'onedrive' && cloud.configured())
			content.innerHTML =
				'<p>Open a connected Office document. Save updates the same file.</p><label>File link or reference<input id="onedrive-link" name="link" required autocomplete="off" spellcheck="false" /></label><button class="primary">Open</button>';
		content.innerHTML = html(content.innerHTML);
		content.onsubmit = (event) => {
			event.preventDefault();
			const button = content.querySelector('button');
			button.disabled = true;
			void task(async () => {
				try {
					const doc = await cloud.open(new FormData(content).get('link').trim());
					$('dialog').close();
					await open(doc.id, undefined, { focus: true });
					await refresh();
				} catch {
					notify(
						message('OneDrive could not open this file. Check sign-in, the link and your access.'),
					);
				} finally {
					button.disabled = false;
				}
			});
		};
		choose(message('Open from OneDrive'), content);
	};
	cloud.showShare = (doc) => {
		const content = document.createElement('form');
		content.id = 'onedrive-share-form';
		content.innerHTML = `<p>Share the connected OneDrive file. The existing link keeps current permissions. Creating a new view or edit link grants the access you select; tenant policy may restrict it. This does not upload another Teams attachment or enable Microsoft coauthoring.</p><label>Access<select name="type"><option value="existing">Existing link, keep current permissions</option><option value="view">Create view link</option><option value="edit">Create edit link</option></select></label><label>New link scope<select name="scope"><option value="organization">People in your organization</option><option value="anonymous">Anyone with the link</option></select></label><button class="primary">Get OneDrive link</button><label>Link<input id="onedrive-share-link" readonly /></label>`;
		const choices =
			options.shareChoices ??
			(options.id === 'onedrive' && options.client?.share
				? null
				: [{ label: 'Existing link, keep permissions', options: { type: 'existing' } }]);
		if (choices) {
			content.innerHTML = `<p>Choose who can access this file.</p><label>Access<select name="choice">${choices.map((choice, index) => `<option value="${index}">${escape(choice.label)}</option>`).join('')}</select></label><button class="primary">Get link</button><label>Link<input id="onedrive-share-link" readonly /></label>`;
			content.querySelector('button').disabled = choices.length === 0;
		}
		content.innerHTML = html(content.innerHTML);
		content.onsubmit = (event) => {
			event.preventDefault();
			const button = content.querySelector('button');
			button.disabled = true;
			void task(async () => {
				try {
					const data = new FormData(content);
					content.querySelector('#onedrive-share-link').value = await cloud.share(
						await store.get(doc.id),
						choices
							? choices[Number(data.get('choice'))].options
							: { type: data.get('type'), scope: data.get('scope') },
					);
					notify(message('OneDrive link ready. Copy it to share with the intended recipients.'));
				} catch {
					notify(
						message(
							'The OneDrive link could not be created. Check sign-in, sharing permissions and organization policy.',
						),
					);
				} finally {
					button.disabled = false;
				}
			});
		};
		choose(message('Share in OneDrive'), content);
	};
	const browser = options.fileProvider
		? mountCloudBrowser({ cloud, options, open, refresh })
		: undefined;
	/** The browsing view when the host can list files; otherwise the link dialog (or its notice). */
	cloud.showFiles = async () => {
		if (!browser) return cloud.showOpen();
		await showView();
		await browser.show();
	};
	cloud.hasBrowser = Boolean(browser);
	$('open-onedrive').textContent = message(browser ? 'Browse OneDrive' : 'Open from OneDrive');
	$('open-onedrive').onclick = () => void task(cloud.showFiles);
	return cloud;
}
