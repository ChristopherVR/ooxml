import { $, escape, task, notify, choose, download } from './ui.js';
import { product } from './product.js';
import {
	desktop,
	native,
	forgetSnapshot,
	chooseFolder,
	entries,
	readEntry,
	locationRecords,
} from './disk-access.js';
const officeKind = (name) => /\.(docx|xlsx|pptx|vsdx)$/i.exec(name)?.[1].toLowerCase();
export function mountDiskSearch({ store, open, refresh }) {
	let folders = [],
		controller,
		results = [],
		lastFilter = '';
	const panel = $('disk-search'),
		menu = $('storage-menu');
	function cancel() {
		controller?.abort();
		panel.removeAttribute('aria-busy');
	}
	async function connect(replaceId) {
		try {
			const folder = await chooseFolder();
			if (!folder) return;
			if (typeof replaceId === 'string' && replaceId !== folder.id) {
				await locationRecords('folders', 'delete', replaceId);
				forgetSnapshot(replaceId);
				folders = folders.filter((f) => f.id !== replaceId);
			}
			if (folder.type !== 'snapshot') await locationRecords('folders', 'put', folder);
			folders = folders.filter((f) => f.id !== folder.id);
			folders.push(folder);
			renderStorage();
			search(lastFilter);
		} catch (error) {
			if (error.name !== 'AbortError') notify(error);
		}
	}
	function renderStorage() {
		menu.innerHTML = `<div class="storage-menu-heading"><h2>Storage & folders</h2><p>Your app library stores editable copies. Connect folders to find files on disk by name.</p></div><button data-connect-folder>Connect folder…</button><button data-refresh-folders>Refresh folder search</button><div class="connected-folders">${folders.map((f) => `<div><strong>${escape(f.name)}</strong><small>${f.type === 'native' ? 'Desktop folder' : f.type === 'snapshot' ? 'Selected files, this session only' : 'Connected folder'}</small><button data-reconnect="${escape(f.id)}">Reconnect</button><button data-disconnect="${escape(f.id)}">Disconnect</button></div>`).join('') || '<p>No folders connected.</p>'}</div><p class="storage-note">${desktop ? 'Files are searched only inside folders you choose.' : window.showDirectoryPicker ? 'Your browser may ask to renew folder access after reopening.' : 'This browser can search a selected folder snapshot. Select it again to refresh; access ends when this page closes.'}</p>`;
		menu.querySelector('[data-connect-folder]').onclick = connect;
		menu.querySelector('[data-refresh-folders]').onclick = () => search(lastFilter);
		menu.querySelectorAll('[data-disconnect]').forEach(
			(b) =>
				(b.onclick = () =>
					task(async () => {
						await locationRecords('folders', 'delete', b.dataset.disconnect);
						forgetSnapshot(b.dataset.disconnect);
						folders = folders.filter((f) => f.id !== b.dataset.disconnect);
						renderStorage();
						search(lastFilter);
					})),
		);
		menu.querySelectorAll('[data-reconnect]').forEach(
			(b) =>
				(b.onclick = async () => {
					const folder = folders.find((f) => f.id === b.dataset.reconnect);
					try {
						if (folder.handle) {
							await folder.handle.requestPermission({ mode: 'read' });
							search(lastFilter);
						} else await connect(folder.id);
					} catch (error) {
						notify(error);
					}
				}),
		);
	}
	async function search(filter = '') {
		cancel();
		lastFilter = filter;
		const signal = (controller = new AbortController()).signal;
		const query = $('search').value.trim().toLowerCase();
		panel.hidden = !query && !folders.length;
		if (panel.hidden) return;
		panel.setAttribute('aria-busy', 'true');
		panel.innerHTML = '<p class="search-progress">Searching connected folders…</p>';
		const found = [],
			issues = [];
		try {
			for (const folder of folders) {
				delete folder.warning;
				try {
					for await (const entry of entries(folder, signal)) {
						signal.throwIfAborted();
						const kind = officeKind(entry.name);
						if (
							entry.path.toLowerCase().includes(query) &&
							(!product || product === 'teams' || kind === product) &&
							(!['docx', 'xlsx', 'pptx', 'vsdx'].includes(filter) || kind === filter) &&
							filter !== 'favourites'
						)
							found.push({ ...entry, folderName: folder.name });
					}
					if (folder.warning) issues.push(`${folder.name}: ${folder.warning}`);
				} catch (error) {
					if (signal.aborted) return;
					issues.push(
						`${folder.name}: ${error.name === 'NotAllowedError' ? 'Reconnect this folder to renew access.' : error.message}`,
					);
				}
			}
			signal.throwIfAborted();
			results = found.slice(0, 100);
			panel.innerHTML = `<div class="disk-results-heading"><h3>Connected folders</h3><span>${found.length} matches${found.length > 100 ? ' (first 100 shown)' : ''}</span><button data-connect-folder>Connect folder…</button></div>${issues.map((message) => `<p class="search-issue">${escape(message)} <button data-storage>Manage folders</button></p>`).join('')}${!folders.length ? '<p>Connect a folder to include files on disk in your search.</p>' : ''}<ul>${results.map((entry, i) => `<li><button data-disk-open="${i}"><strong>${escape(entry.name)}</strong><small>${escape(entry.folderName + '/' + entry.path)}</small></button><button data-disk-location="${i}" aria-label="File location: ${escape(entry.name)}">Location</button></li>`).join('')}</ul>`;
			panel.querySelector('[data-connect-folder]').onclick = connect;
		} catch (error) {
			if (!signal.aborted) notify(error);
		} finally {
			if (!signal.aborted) panel.removeAttribute('aria-busy');
		}
	}
	async function location(source) {
		const folder = folders.find((f) => f.id === source.folderId);
		const content = document.createElement('div');
		content.innerHTML = `<p class="file-location">${escape(source.folderName + '/' + source.path)}</p><p>Editing saves the workspace copy. Use Download to export your changes.</p>`;
		if (folder?.type === 'native') {
			const button = document.createElement('button');
			button.textContent = navigator.userAgent.includes('Mac')
				? 'Show in Finder'
				: 'Show in Explorer';
			button.onclick = () =>
				task(() => native('disk_reveal', { root: folder.root, path: source.path }));
			content.append(button);
		} else
			content.insertAdjacentHTML(
				'beforeend',
				'<p>This browser does not expose an absolute disk path or a reveal-in-file-manager action.</p>',
			);
		choose('File location', content);
	}
	async function locate(id) {
		const source = await locationRecords('sources', 'get', id);
		if (source) return location(source);
		const content = document.createElement('p');
		content.textContent =
			'This copy is in app storage for the current local workspace. Files imported with a file picker do not expose their original disk location. Connect the containing folder to search it, or download a copy.';
		choose('File location', content);
	}
	panel.addEventListener('click', (event) => {
		const button = event.target.closest('[data-disk-open], [data-disk-location]');
		if (!button) return;
		const entry = results[Number(button.dataset.diskOpen ?? button.dataset.diskLocation)];
		if (button.hasAttribute('data-disk-location')) return task(() => location(entry));
		task(async () => {
			const folder = folders.find((f) => f.id === entry.folderId);
			const file = await readEntry(folder, entry);
			if (!officeKind(file.name)) {
				download(file);
				return;
			}
			const sources = await locationRecords('sources', 'getAll');
			const docs = await store.list();
			const previous = sources.find(
				(s) =>
					s.folderId === entry.folderId && s.path === entry.path && docs.some((d) => d.id === s.id),
			);
			const doc = previous
				? await store.get(previous.id)
				: await store.create(file.name, file.bytes);
			if (!previous)
				await locationRecords('sources', 'put', {
					id: doc.id,
					folderId: entry.folderId,
					folderName: folder.name,
					path: entry.path,
				});
			await open(doc.id);
			await refresh();
		});
	});
	document.addEventListener('click', (event) => {
		if (event.target.closest('[data-storage]')) menu.showPopover();
	});
	task(async () => {
		folders = await locationRecords('folders', 'getAll');
		renderStorage();
		if (folders.length) search();
	});
	return { search, cancel, locate };
}
