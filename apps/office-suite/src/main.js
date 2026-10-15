import { SuiteDocumentStore } from 'ooxml-ui/suite';
import { listEmbeddedPackages } from 'ooxml-core/opc';
import { SuiteSessions } from './session.js';
import { showVersions, withVersions } from './versions.js';
import { currentPolicy } from './ai/policy.js';
import { blankDocument, preloadEditor } from './editors.js';
import { mountTeams, shareDocument, localDocumentId } from './teams.js';
import { $, apps, badge, escape, notify, task, download, choose, tag, teamsApp } from './ui.js';
import { mountThemes } from './themes.js';
import { mountProfile } from './profile.js';
import { documentDatabase } from './profile-state.js';
import { renderLibrary, toggleFavourite } from './home.js';
import { product, mountProduct } from './product.js';
import { mountFileActions } from './file-actions.js';
import { mountDiskSearch } from './disk-search.js';
import { mountPwa } from './pwa.js';
import { mountResources, mountStart } from './start.js';
import { mountShell } from './shell.js';
import { icon } from './icons.js';
import { localAttachmentId, showAttachment } from './attachments.js';

// Every write keeps the bytes it replaces as a version, unless the AI policy turns history off.
const store = withVersions(
	new SuiteDocumentStore(documentDatabase()),
	() => currentPolicy().features.versions,
);
let files = [],
	active = null,
	view = 'home',
	filter = '',
	teams,
	teamsTabOpen = false,
	chatTabOpen = false,
	chatPage,
	assistant;
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
	// The dialog offers the blank documents only; Teams and the file picker stay on the page.
	for (const button of content.querySelectorAll('button:not([data-create])')) button.remove();
	content.addEventListener('click', (event) => {
		if (event.target.closest('[data-create]')) $('dialog').close();
	});
	choose('Create something new', content);
}
let applyShell = () => {};
function renderTabs() {
	document.body.dataset.view = view;
	$('tabs').hidden = !teamsTabOpen && !chatTabOpen && !sessions.sessions.size;
	$('tabs').innerHTML =
		(chatTabOpen
			? `<button role="tab" aria-selected="${view === 'chat'}" data-chat>${icon('spark')}Chat</button>`
			: '') +
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
	$('chat').hidden = view !== 'chat';
	$('teams-host').hidden = view !== 'teams';
	$('editor-area').hidden = view !== 'editor';
	$('back-teams').hidden = !origins.has(active);
	$('back-parent').hidden = !sessions.sessions.get(active)?.doc.parent;
	document.title =
		active && view === 'editor'
			? `${sessions.sessions.get(active)?.doc.name ?? 'Document'} - OOXML Office`
			: view === 'chat'
				? 'Chat - OOXML Office'
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
	applyShell();
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
/** What the chat page and the assistant pane need from the shell. */
const aiDeps = {
	store,
	sessions,
	refresh,
	open,
	getTeams: () => teams,
	openIds: () => [...sessions.sessions.keys()],
	showVersions: (id) => task(() => showVersions({ id, store, sessions, refresh })),
};
/** The chat page; `id` reopens a conversation from the history. */
async function showChat(id) {
	chatTabOpen = true;
	view = 'chat';
	renderTabs();
	if (!chatPage) {
		const { mountChatPage } = await import('./ai/chat-page.js');
		chatPage = mountChatPage({
			host: $('chat'),
			deps: aiDeps,
			onRoute: (conversation) => {
				if (view === 'chat')
					history.replaceState(null, '', conversation ? `#/chat/${conversation}` : '#/chat');
			},
		});
	}
	await chatPage.show(id);
	history.replaceState(null, '', chatPage.id ? `#/chat/${chatPage.id}` : '#/chat');
}
const chatRoute = () => /^#\/chat(?:\/([\w-]{1,100}))?$/u.exec(location.hash);
async function showAssistant(show = true) {
	if (!assistant) {
		const { mountAssistant } = await import('./assistant.js');
		assistant = mountAssistant({
			deps: aiDeps,
			getActive: () => (view === 'editor' ? active : null),
			openChat: (id) =>
				task(async () => {
					await showAssistant(false);
					await showChat(id ?? undefined);
				}),
		});
	}
	$('ai-close').onclick = () => task(() => showAssistant(false));
	$('assistant').hidden = !show;
	$('assistant-toggle').setAttribute('aria-expanded', String(show));
	if (show) await assistant.opened();
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
	`<span class="nav-caption">Workspace</span><button data-chat data-view="chat">${icon('spark')}<span>Chat</span></button><button data-home data-view="home" aria-current="true">${icon('home')}<span>Home</span></button><button data-library="files">${icon('folder')}<span>My files</span></button><button data-library="favourites">${icon('star')}<span>Favourites</span></button><span class="nav-caption">Apps</span><button data-teams data-view="teams">${badge('teams')}<span>Teams</span>${tag(teamsApp.tag)}</button>` +
	apps
		.map(
			(a) =>
				`<button data-app="${a.id}">${badge(a.id)}<span>${a.name}</span>${tag(a.tag)}</button>`,
		)
		.join('');
const blankDescriptions = {
	docx: 'Write and format a document',
	xlsx: 'Calculate, chart and sort a workbook',
	pptx: 'Build and present a slide deck',
	vsdx: 'Draw a diagram or flowchart',
};
$('create-row').innerHTML =
	apps
		.map(
			(a) =>
				`<button data-create="${a.id}">${badge(a.id)}<span>${blankNames[a.id]}${tag(a.tag)}</span><small>${blankDescriptions[a.id]} in ${a.name}</small></button>`,
		)
		.join('') +
	`<button data-teams>${badge('teams')}<span>Teams${tag(teamsApp.tag)}</span><small>Chat, calls and shared files on your server</small></button><button data-import>${icon('upload')}<span>Open a file</span><small>Bring a .docx, .xlsx, .pptx or .vsdx from this device</small></button>`;
// Each tile's place in the entrance stagger (styles.css).
for (const [i, button] of [...$('create-row').children].entries())
	button.style.setProperty('--i', i);
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
	if (b.hasAttribute('data-chat')) task(() => showChat());
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
$('versions').onclick = () => aiDeps.showVersions(active);
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
$('assistant-toggle').onclick = () => task(() => showAssistant($('assistant').hidden));
$('ask-ai').onclick = () => task(() => showAssistant(true));
$('home-ask').onsubmit = (event) => {
	event.preventDefault();
	const prompt = $('home-ask-input').value.trim();
	if (!prompt) return;
	$('home-ask-input').value = '';
	task(async () => {
		await showChat();
		chatPage.ask(prompt);
	});
};
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
	// "/" focuses the search from anywhere in the shell, as the hint beside it says.
	if (
		event.key === '/' &&
		!event.ctrlKey &&
		!event.metaKey &&
		!event.altKey &&
		!event.target.closest(
			'input, textarea, [contenteditable], docx-editor, xlsx-editor, visio-viewer, teams-app, .pptxv',
		)
	) {
		event.preventDefault();
		$('search').focus();
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
// The home rises into place once; after that, switching back to it is instant (styles.css).
document.body.classList.add('entering');
setTimeout(() => document.body.classList.remove('entering'), 1200);
mountThemes();
mountProfile({
	beforeSwitch: async () => {
		for (const id of sessions.sessions.keys()) await sessions.save(id);
		teams?.client?.flushStorage();
	},
	changed: () => renderFiles(),
});
$('app-launcher').innerHTML =
	`<div class="launcher-heading"><h2>Your apps</h2><span>OOXML Office</span></div><div class="launcher-grid"><button data-chat><span class="app-badge chat-badge">${icon('spark')}</span><span>Chat</span></button>${apps.map((a) => `<button data-app="${a.id}">${badge(a.id)}<span>${a.name}</span></button>`).join('')}<button data-teams>${badge('teams')}<span>Teams</span></button></div>`;
$('app-launcher').addEventListener('click', (event) => {
	if (event.target.closest('button')) $('app-launcher').hidePopover();
});
$('launcher-toggle').innerHTML = icon('grid');
$('theme-picker').innerHTML = icon('settings');
$('assistant-toggle').innerHTML = `${icon('spark')}<span>Assistant</span>`;
$('ask-ai').innerHTML = `${icon('spark')}<span>Ask AI</span>`;
$('home-ask-symbol').innerHTML = icon('spark');
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
		else if (chatRoute() || (!product && !location.hash)) await showChat(chatRoute()?.[1]);
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
	// The suite opens on chat; a product's own app opens on its files.
	else if (chatRoute() || (!product && !location.hash)) await showChat(chatRoute()?.[1]);
});

for (const eventName of ['pointerover', 'focusin'])
	document.addEventListener(eventName, (event) => {
		const target = event.target.closest('[data-create], [data-app]');
		if (target) void preloadEditor(target.dataset.create ?? target.dataset.app)?.catch(() => {});
	});

mountProduct();
applyShell = mountShell(() => ({
	view,
	documents: sessions.sessions.size,
	teamsOpen: teamsTabOpen,
}));
mountPwa({ idle: () => sessions.sessions.size === 0 && view === 'home' });
mountStart();
mountResources();

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
