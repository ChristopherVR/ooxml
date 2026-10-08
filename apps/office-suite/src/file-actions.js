import { $, task, notify, choose, escape, download } from './ui.js';
import { contextMenu } from './context-menu.js';
export function mountFileActions({
	store,
	sessions,
	refresh,
	closeTeams,
	active,
	activate,
	home,
	locate,
}) {
	let closing = false;
	async function closeTabs(ids) {
		if (closing) return;
		closing = true;
		const ordered = [...sessions.sessions.keys()];
		const previous = active();
		const index = ordered.indexOf(previous);
		const unlock = sessions.lock(ids);
		try {
			// Save all targets before closing any: a failed save leaves the tabs available.
			for (const id of ids) await sessions.save(id);
			for (const id of ids) sessions.close(id);
			if (ids.includes(previous)) {
				const remaining = ordered.filter((id) => !ids.includes(id));
				const next = remaining[Math.min(index, remaining.length - 1)];
				if (next) await activate(next);
				else home();
			}
			await refresh();
		} finally {
			unlock();
			closing = false;
		}
	}
	async function remove(id) {
		const doc = await store.get(id);
		const content = document.createElement('div');
		content.innerHTML = `<p>Remove <strong>${escape(doc.name)}</strong> from this workspace? The original file on disk is kept. Workspace links to this copy will stop opening until you undo removal.</p><button class="primary" data-confirm-remove>Remove from workspace</button><button data-cancel-remove>Cancel</button>`;
		choose('Remove file', content);
		content.querySelector('[data-cancel-remove]').onclick = () => $('dialog').close();
		content.querySelector('[data-confirm-remove]').onclick = () =>
			task(async () => {
				const all = await store.list();
				const ids = new Set([id]);
				let count = 0;
				while (count !== ids.size) {
					count = ids.size;
					for (const d of all) if (d.parent && ids.has(d.parent.id)) ids.add(d.id);
				}
				const unlock = sessions.lock([...ids]);
				try {
					// Dirty embedded documents save first, so their parent snapshots contain the edits.
					for (const child of [...ids].reverse()) await sessions.save(child);
					const current = await store.get(id);
					const removed = await store.remove(id, current.revision);
					for (const child of ids) sessions.close(child);
					if (ids.has(active())) home();
					await refresh();
					$('dialog').close();
					notify(`${doc.name} removed. Original disk file kept.`);
					document.querySelector('.file-undo')?.remove();
					const undo = document.createElement('button');
					undo.className = 'file-undo';
					undo.textContent = 'Undo removal';
					undo.onclick = () =>
						task(async () => {
							await store.commit(removed.map((document) => ({ document, expected: 0 })));
							await refresh();
							undo.remove();
							notify('File restored');
						});
					$('status').after(undo);
				} finally {
					unlock();
				}
			});
	}
	function fileMenu(event, id, tab = false) {
		const ids = [...sessions.sessions.keys()],
			index = ids.indexOf(id);
		const item = (label, run, disabled = false) => ({ label, run, disabled, error: notify });
		contextMenu(event, [
			...(tab
				? [
						item('Close', () => closeTabs([id])),
						item('Close others', () => closeTabs(ids.filter((key) => key !== id)), ids.length < 2),
						item(
							'Close tabs to the right',
							() => closeTabs(ids.slice(index + 1)),
							index === ids.length - 1,
						),
						item('Close all documents', () => closeTabs(ids)),
					]
				: []),
			item('Download a copy', async () => download(await sessions.save(id))),
			item('File location', () => locate(id)),
			item('Remove from workspace…', () => remove(id)),
		]);
	}
	$('tabs').addEventListener('mousedown', (event) => {
		if (event.button === 1 && event.target.closest('.file-tab')) event.preventDefault();
	});
	$('tabs').addEventListener('auxclick', (event) => {
		if (event.button === 1 && event.target.closest('[data-teams]')) {
			event.preventDefault();
			closeTeams();
			return;
		}
		const tab = event.target.closest('.file-tab');
		if (event.button !== 1 || !tab) return;
		event.preventDefault();
		task(() => closeTabs([tab.querySelector('[data-open]').dataset.open]));
	});
	$('tabs').addEventListener('contextmenu', (event) => {
		if (event.target.closest('[data-teams]')) {
			contextMenu(event, [{ label: 'Close', run: closeTeams, error: notify }]);
			return;
		}
		const tab = event.target.closest('.file-tab');
		if (tab) fileMenu(event, tab.querySelector('[data-open]').dataset.open, true);
	});
	$('files').addEventListener('contextmenu', (event) => {
		const row = event.target.closest('tr');
		if (row) fileMenu(event, row.querySelector('[data-open]').dataset.open);
	});
	document.addEventListener('click', (event) => {
		const b = event.target.closest('[data-file-menu]');
		if (b) fileMenu(event, b.dataset.fileMenu);
	});
	return { closeTabs };
}
