import { createOneDriveClient, createOneDriveSession } from 'ooxml-core/embedding/onedrive';
import { SuiteOneDrive } from './cloud-files.js';
import { mergeOneDriveChanges } from './onedrive-merge.js';
import { $, choose, download, escape, notify, task } from './ui.js';

/** Shared shell controls for Word, Excel, PowerPoint and Visio, including standalone apps. */
export function mountOneDrive({ store, sessions, open, refresh }) {
	const options = globalThis.ooxmlOneDriveOptions;
	const cloud = new SuiteOneDrive({
		store,
		client: options ? createOneDriveClient(options) : undefined,
		createSession: createOneDriveSession,
		merge: mergeOneDriveChanges,
		changed: () => {
			void refresh();
		},
		adopted: (doc) => sessions.reload(doc),
		dirty: (id) => sessions.isDirty(id),
		conflict: ({ doc, conflict }) =>
			new Promise((resolve) => {
				const content = document.createElement('div');
				content.id = 'onedrive-conflict';
				content.innerHTML = `<p>Some edits in <strong>${escape(doc.name)}</strong> could not be combined automatically. Independent changes are merged during Save; the remaining changes overlap or cannot be matched safely. Your edits are saved on this device.</p>
				<p>The OneDrive version is ${escape(conflict.remote.file.name)}. Saving a copy keeps the original unchanged. Opening the latest first keeps your edits as a separate file in your library.</p>
				<div class="cloud-recovery-actions"><button data-cloud-recovery="copy" class="primary">Save my edits as a OneDrive copy</button><button data-cloud-recovery="latest">Keep local copy and open latest</button><button data-cloud-recovery="download">Download my edits</button><button data-cloud-recovery="download-latest">Download latest OneDrive version</button><button data-cloud-recovery="cancel">Keep editing</button></div>`;
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
			}),
	});
	cloud.showOpen = () => {
		const content = document.createElement('form');
		content.id = 'onedrive-open-form';
		content.innerHTML = cloud.configured()
			? '<p>Open a Word, Excel, PowerPoint or Visio file using your Microsoft access. Save updates the connected file. Live Microsoft coauthoring is not enabled.</p><label>OneDrive sharing link<input id="onedrive-link" name="link" type="url" required autocomplete="off" spellcheck="false" /></label><button id="onedrive-open-submit" class="primary">Open from OneDrive</button>'
			: '<p>Microsoft sign-in has not been connected by your host or administrator. Ask them to enable OneDrive access. You can continue opening files from this device.</p><p><a href="https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#integrated-office-suite" target="_blank" rel="noopener noreferrer">Integration documentation</a></p>';
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
					notify('OneDrive could not open this file. Check sign-in, the link and your access.');
				} finally {
					button.disabled = false;
				}
			});
		};
		choose('Open from OneDrive', content);
	};
	cloud.showShare = (doc) => {
		const content = document.createElement('form');
		content.id = 'onedrive-share-form';
		content.innerHTML = `<p>Share the connected OneDrive file. The existing link keeps current permissions. Creating a new view or edit link grants the access you select; tenant policy may restrict it. This does not upload another Teams attachment or enable Microsoft coauthoring.</p><label>Access<select name="type"><option value="existing">Existing link, keep current permissions</option><option value="view">Create view link</option><option value="edit">Create edit link</option></select></label><label>New link scope<select name="scope"><option value="organization">People in your organization</option><option value="anonymous">Anyone with the link</option></select></label><button class="primary">Get OneDrive link</button><label>Link<input id="onedrive-share-link" readonly /></label>`;
		content.onsubmit = (event) => {
			event.preventDefault();
			const button = content.querySelector('button');
			button.disabled = true;
			void task(async () => {
				try {
					const data = new FormData(content);
					content.querySelector('#onedrive-share-link').value = await cloud.share(
						await store.get(doc.id),
						{ type: data.get('type'), scope: data.get('scope') },
					);
					notify('OneDrive link ready. Copy it to share with the intended recipients.');
				} catch {
					notify(
						'The OneDrive link could not be created. Check sign-in, sharing permissions and organization policy.',
					);
				} finally {
					button.disabled = false;
				}
			});
		};
		choose('Share in OneDrive', content);
	};
	$('open-onedrive').onclick = cloud.showOpen;
	return cloud;
}
