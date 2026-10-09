import { SuiteDocumentStore } from 'ooxml-ui/suite';
import { listEmbeddedPackages } from 'ooxml-core/opc';
import { SuiteSessions } from './session.js';
import { blankDocument, preloadEditor } from './editors.js';
import { mountTeams, shareDocument, localDocumentId } from './teams.js';
import { $, apps, badge, escape, notify, task, download, choose } from './ui.js';
import { mountThemes } from './themes.js';
import { mountProfile } from './profile.js';
import { documentDatabase } from './profile-state.js';
import { renderLibrary, toggleFavourite } from './home.js';
import { product, mountProduct } from './product.js';
import { mountFileActions } from './file-actions.js';
import { mountDiskSearch } from './disk-search.js';
import { mountPwa } from './pwa.js';
import { mountStart } from './start.js';
import { icon } from './icons.js';
import { localAttachmentId, showAttachment } from './attachments.js';

const store = new SuiteDocumentStore(documentDatabase());
let files = [],
	active = null,
	view = 'home',
	filter = '',
	teams,
	teamsTabOpen = false;
const origins = new Map();
let disk;
const sessions = new SuiteSessions(store, $('editors'), renderTabs, (command, id) =>
	task(async () => {
		if (command === 'save') {
			await sessions.save(id);
			await refresh();
			notify('Saved to this device');
		} else if (command === 'open') $('file-input').click();
		else await create((await store.get(id)).kind);
	}),
);

async function refresh() {
	files = await store.list();
	renderFiles();
	renderTabs();
}
/** The app the home page is about (a rail filter or a standalone product page), if any. */
function createKind() {
	return apps.find((a) => a.id === filter || a.id === product)?.id;
}
function renderFiles() {
	renderLibrary(files, filter);
	disk?.search(filter);
	const kind = createKind();
	$('create-new').hidden = product === 'teams';
	$('create-new').innerHTML =
		`${icon('add')}<span>${kind ? `New ${blankNames[kind].toLowerCase()}` : 'Create'}</span>`;
}
function chooseBlank() {
	const content = document.createElement('div');
	content.className = 'create-row';
	content.innerHTML = $('create-row').innerHTML;
	content.addEventListener('click', (event) => {
		if (event.target.closest('[data-create]')) $('dialog').close();
	});
	choose('Create something new', content);
}
function renderTabs() {
	document.body.dataset.view = view;
	$('tabs').hidden = !teamsTabOpen && !sessions.sessions.size;
	$('tabs').innerHTML =
		`<button role="tab" aria-selected="${view === 'home'}" data-home>Home</button>` +
		(teamsTabOpen
			? `<button role="tab" aria-selected="${view === 'teams'}" data-teams>Teams</button>`
			: '') +
		[...sessions.sessions.values()]
			.map(
				(s) =>
					`<span class="file-tab"><button role="tab" aria-selected="${s.doc.id === active && view === 'editor'}" data-open="${s.doc.id}">${badge(s.doc.kind)}${escape(s.doc.name)}${sessions.isDirty(s.doc.id) ? ' •' : ''}</button><button data-close="${s.doc.id}" aria-label="Close ${escape(s.doc.name)}">×</button></span>`,
			)
			.join('');
	$('save-state').textContent =
		active && sessions.isDirty(active) ? 'Unsaved changes' : 'Saved to this device';
	for (const s of sessions.sessions.values())
		s.node.hidden = view !== 'editor' || s.doc.id !== active;
	$('home').hidden = view !== 'home';
	$('teams-host').hidden = view !== 'teams';
	$('editor-area').hidden = view !== 'editor';
	$('back-teams').hidden = !origins.has(active);
	$('back-parent').hidden = !sessions.sessions.get(active)?.doc.parent;
	document.title =
		active && view === 'editor'
			? `${sessions.sessions.get(active)?.doc.name ?? 'Document'} - OOXML Office`
			: product
				? `OOXML ${apps.find((a) => a.id === product)?.name ?? 'Teams'}`
				: 'OOXML Office';
	for (const button of $('rail').querySelectorAll('button'))
		button.setAttribute(
			'aria-current',
			String(
				(button.dataset.view === view && (view !== 'home' || !filter)) ||
					(view === 'home' &&
						!!filter &&
						(button.dataset.app === filter || button.dataset.library === filter)) ||
					(view === 'editor' && button.dataset.app === sessions.sessions.get(active)?.doc.kind),
			),
		);
}
async function open(id, origin) {
	if (origin) origins.set(id, origin);
	const doc = await store.get(id);
	active = id;
	view = 'editor';
	renderTabs();
	notify(`Opening ${doc.name}...`);
	await sessions.open(doc);
	renderTabs();
	await refresh();
	notify(
		doc.parent
			? 'Embedded document: Save updates its parent package.'
			: origin?.remote
				? 'Local copy of a Teams attachment. Share in Teams uploads a new copy.'
				: 'Ready',
	);
	history.replaceState(null, '', `#/file/${id}`);
}
async function showTeams() {
	teamsTabOpen = true;
	view = 'teams';
	renderTabs();
	if (!teams) teams = await mountTeams($('teams-host'), store, open, refresh);
	renderTabs();
	history.replaceState(null, '', '#/teams');
}
const blankNames = {
	docx: 'Document',
	xlsx: 'Workbook',
	pptx: 'Presentation',
	vsdx: 'Drawing',
};
async function create(kind) {
	void preloadEditor(kind)?.catch(() => {});
	const name = `${blankNames[kind]} ${files.filter((d) => d.kind === kind).length + 1}.${kind}`;
	const doc = await store.create(name, await blankDocument(kind));
	await open(doc.id);
}
async function importFiles(items) {
	for (const file of items) {
		if (file.size > 32 * 1024 * 1024) throw new Error('Choose a file smaller than 32 MB.');
		const doc = await store.create(file.name, new Uint8Array(await file.arrayBuffer()));
		await open(doc.id);
	}
}
async function embeddings() {
	const doc = await sessions.save(active);
	const embedded = await listEmbeddedPackages(doc.bytes);
	const content = document.createElement('div');
	content.innerHTML =
		'<p>Open an embedded document in its editor. Saving writes it back into this package. Linked chart previews may require refresh in the originating Office application.</p>';
	if (!embedded.length)
		content.insertAdjacentHTML('beforeend', '<p>No embedded Office packages in this file.</p>');
	for (const item of embedded) {
		const button = document.createElement('button');
		button.className = 'embedded-file';
		button.textContent = `${item.name}${item.kind === 'unsupported' ? ' (legacy OLE: download only)' : ''}`;
		button.onclick = () =>
			task(async () => {
				if (item.kind === 'unsupported') {
					download({ name: item.name, bytes: item.bytes });
					return;
				}
				const existing = (await store.list()).find(
					(d) =>
						d.parent?.id === doc.id &&
						d.parent.path === item.path &&
						d.parent.revision === doc.revision,
				);
				const child =
					existing ??
					(await store.create(item.name, item.bytes, {
						parent: { id: doc.id, path: item.path, revision: doc.revision },
					}));
				$('dialog').close();
				await open(child.id);
			});
		content.append(button);
	}
	choose('Embedded documents', content);
}
async function share() {
	const doc = await sessions.save(active);
	if (!teams) {
		teams = await mountTeams($('teams-host'), store, open, refresh);
	}
	const client = teams.client;
	if (!client) throw new Error('Teams is starting. Please try again.');
	const content = document.createElement('div');
	content.innerHTML =
		client.getState().mode === 'local'
			? '<p>Share a link to this file in a local Teams channel. Everyone using this browser workspace opens the same saved document.</p>'
			: '<p>Upload the edited document as a new copy in this server channel. The original attachment is preserved.</p>';
	for (const channel of client.getState().channels) {
		const b = document.createElement('button');
		b.className = 'embedded-file';
		b.textContent = channel.name;
		b.onclick = () =>
			task(async () => {
				await shareDocument(teams, doc, channel.id);
				$('dialog').close();
				client.select(channel.id);
				await showTeams();
				notify(`Shared ${doc.name} in ${channel.name}`);
			});
		content.append(b);
	}
	choose(`Share ${doc.name}`, content);
}

$('rail').innerHTML =
	`<button data-home data-view="home" aria-current="true">${icon('home')}<span>Home</span></button><button data-library="files">${icon('folder')}<span>My files</span></button><button data-library="favourites">${icon('star')}<span>Favourites</span></button><span class="nav-caption">APPLICATIONS</span><button data-teams data-view="teams">${badge('teams')}<span>Teams</span></button>` +
	apps
		.map((a) => `<button data-app="${a.id}">${badge(a.id)}<span>${a.name}</span></button>`)
		.join('');
$('create-row').innerHTML = apps
	.map(
		(a) =>
			`<button data-create="${a.id}">${badge(a.id)}<span>Blank ${blankNames[a.id].toLowerCase()}</span><small>${a.name}</small></button>`,
	)
	.join('');
document.addEventListener('click', (event) => {
	const b = event.target.closest('button');
	if (!b) return;
	if (b.hasAttribute('data-home')) {
		view = 'home';
		filter = '';
		renderTabs();
		renderFiles();
		history.replaceState(null, '', '#/');
	}
	if (b.hasAttribute('data-import')) $('file-input').click();
	if (b.hasAttribute('data-teams')) task(showTeams);
	if (b.dataset.app) {
		view = 'home';
		filter = b.dataset.app;
		renderTabs();
		renderFiles();
	}
	if (b.hasAttribute('data-library') || b.hasAttribute('data-file-filter')) {
		view = 'home';
		filter = b.dataset.library ?? b.dataset.fileFilter;
		renderTabs();
		renderFiles();
	}
	if (b.dataset.favourite) {
		toggleFavourite(b.dataset.favourite);
		renderFiles();
	}
	if (b.dataset.create) task(() => create(b.dataset.create));
	if (b.dataset.open) task(() => open(b.dataset.open));
	if (b.dataset.download) task(async () => download(await sessions.save(b.dataset.download)));
	if (b.dataset.close) task(() => fileActions.closeTabs([b.dataset.close]));
});
$('upload').onclick = () => $('file-input').click();
$('create-new').onclick = () => task(() => (createKind() ? create(createKind()) : chooseBlank()));
$('file-input').onchange = () =>
	task(async () => {
		const items = [...$('file-input').files];
		$('file-input').value = '';
		await importFiles(items);
	});
let searchTimer;
$('search').oninput = () => {
	view = 'home';
	renderTabs();
	clearTimeout(searchTimer);
	disk.cancel();
	$('files').setAttribute('aria-busy', 'true');
	searchTimer = setTimeout(() => {
		renderFiles();
		$('files').removeAttribute('aria-busy');
	}, 180);
};
$('save').onclick = () =>
	task(async () => {
		await sessions.save(active);
		await refresh();
		notify('Saved to this device');
	});
$('download').onclick = () => task(async () => download(await sessions.save(active)));
$('embedded').onclick = () => task(embeddings);
$('share').onclick = () => task(share);
$('back-teams').onclick = () =>
	task(async () => {
		const origin = origins.get(active);
		if (sessions.isDirty(active)) await sessions.save(active);
		await showTeams();
		if (origin?.channelId) teams.client?.select(origin.channelId);
	});
$('back-parent').onclick = () =>
	task(async () => {
		const doc = await sessions.save(active);
		await open(doc.parent.id);
	});
$('rename').onclick = () =>
	task(async () => {
		const doc = await sessions.save(active);
		const content = document.createElement('form');
		content.innerHTML = `<label>File name<input name="name" value="${escape(doc.name)}" required /></label><button class="primary">Rename</button>`;
		content.onsubmit = (event) => {
			event.preventDefault();
			task(async () => {
				const name = new FormData(content).get('name').trim();
				if (!name.endsWith(`.${doc.kind}`)) throw new Error(`Keep the .${doc.kind} extension.`);
				const next = { ...doc, name, revision: doc.revision + 1 };
				await store.commit([{ document: next, expected: doc.revision }]);
				sessions.sessions.get(doc.id).doc = next;
				$('dialog').close();
				await refresh();
			});
		};
		choose('Rename file', content);
	});
$('dialog-close').onclick = () => $('dialog').close();
let assistantReady = false;
$('assistant-toggle').onclick = () =>
	task(async () => {
		if (!assistantReady) {
			const { mountAssistant } = await import('./assistant.js');
			mountAssistant({ store, sessions, refresh, getActive: () => active, getTeams: () => teams });
			assistantReady = true;
		}
		$('assistant').hidden = !$('assistant').hidden;
		$('assistant-toggle').setAttribute('aria-expanded', String(!$('assistant').hidden));
	});
window.addEventListener('beforeunload', (event) => {
	if ([...sessions.sessions.keys()].some((id) => sessions.isDirty(id))) {
		event.preventDefault();
		event.returnValue = '';
	}
});
document.addEventListener('keydown', (event) => {
	if ((event.ctrlKey || event.metaKey) && event.key === 's' && active) {
		event.preventDefault();
		$('save').click();
	}
});
document.addEventListener('dragover', (event) => {
	if (event.dataTransfer?.types.includes('Files')) event.preventDefault();
});
document.addEventListener('drop', (event) => {
	if (view !== 'teams' && event.dataTransfer?.files.length) {
		event.preventDefault();
		task(() => importFiles([...event.dataTransfer.files]));
	}
});
mountThemes();
mountProfile({
	beforeSwitch: async () => {
		for (const id of sessions.sessions.keys()) await sessions.save(id);
		teams?.client?.flushStorage();
	},
	changed: () => renderFiles(),
});
$('app-launcher').innerHTML =
	`<div class="launcher-heading"><h2>Your apps</h2><span>OOXML Office</span></div><div class="launcher-grid">${apps.map((a) => `<button data-app="${a.id}">${badge(a.id)}<span>${a.name}</span></button>`).join('')}<button data-teams>${badge('teams')}<span>Teams</span></button></div>`;
$('app-launcher').addEventListener('click', (event) => {
	if (event.target.closest('button')) $('app-launcher').hidePopover();
});
$('launcher-toggle').innerHTML = icon('grid');
$('theme-picker').innerHTML = icon('settings');
$('assistant-toggle').innerHTML = `${icon('spark')}<span>Assistant</span>`;
$('search-symbol').innerHTML = icon('search');
$('mobile-search-toggle').innerHTML = icon('search');
$('mobile-search-toggle').onclick = () => {
	const expanded = document.body.classList.toggle('search-expanded');
	$('mobile-search-toggle').setAttribute('aria-expanded', String(expanded));
	if (expanded) $('search').focus();
};
$('upload').innerHTML = `${icon('upload')}<span>Upload file</span>`;
window.addEventListener('hashchange', () =>
	task(async () => {
		const id = localDocumentId(location.href);
		if (id) await open(id);
		else if (localAttachmentId(location.href))
			await showAttachment(localAttachmentId(location.href));
		else if (location.hash === '#/teams' || (product === 'teams' && !location.hash))
			await showTeams();
		else {
			active = null;
			view = 'home';
			filter = '';
			renderTabs();
			renderFiles();
		}
	}),
);
$('tabs').addEventListener('keydown', (event) => {
	if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
	const tabs = [...$('tabs').querySelectorAll('[role=tab]')];
	const index = tabs.indexOf(document.activeElement);
	if (index < 0) return;
	const next =
		event.key === 'Home'
			? 0
			: event.key === 'End'
				? tabs.length - 1
				: (index + (event.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length;
	event.preventDefault();
	tabs[next].focus();
	tabs[next].click();
});
task(async () => {
	await refresh();
	const id = localDocumentId(location.href);
	if (id) await open(id);
	else if (localAttachmentId(location.href)) await showAttachment(localAttachmentId(location.href));
	else if (location.hash === '#/teams' || (product === 'teams' && !location.hash))
		await showTeams();
});

for (const eventName of ['pointerover', 'focusin'])
	document.addEventListener(eventName, (event) => {
		const target = event.target.closest('[data-create], [data-app]');
		if (target) void preloadEditor(target.dataset.create ?? target.dataset.app)?.catch(() => {});
	});

mountProduct();
mountPwa();
mountStart();

disk = mountDiskSearch({ store, open, refresh });
const fileActions = mountFileActions({
	store,
	sessions,
	refresh,
	closeTeams: () => {
		teamsTabOpen = false;
		if (view === 'teams') {
			view = 'home';
			history.replaceState(null, '', '#/');
		}
		renderTabs();
	},
	active: () => active,
	activate: open,
	home: () => {
		active = null;
		view = 'home';
		history.replaceState(null, '', '#/');
		renderTabs();
	},
	locate: (id) => disk.locate(id),
});
