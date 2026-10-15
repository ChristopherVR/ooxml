/**
 * The chat page: the suite's AI-first front door. History on the left, one conversation in the
 * middle, and an empty state that greets the user with starters for each app.
 */
import { $, badge, escape, task, choose } from '../ui.js';
import { icon } from '../icons.js';
import { currentProfile } from '../profile-state.js';
import { conversations, groupByDate } from './conversations.js';
import { createChat } from './chat.js';
import { feature, loadPolicy } from './policy.js';

const starters = [
	{
		kind: 'docx',
		title: 'Draft a project brief',
		prompt:
			'Draft a one-page project brief for redesigning our company website, as a Word document.',
	},
	{
		kind: 'xlsx',
		title: 'Plan a budget',
		prompt:
			'Make a budget workbook for a two-day team offsite for 40 people, with a total formula.',
	},
	{
		kind: 'pptx',
		title: 'Outline a deck',
		prompt: 'Create a five-slide presentation about our product launch plan.',
	},
	{
		kind: 'files',
		title: 'Ask about my files',
		prompt: 'Summarize the attached files and list any open questions.',
	},
];

function greeting() {
	const hour = new Date().getHours();
	const part = hour < 12 ? 'Good morning' : hour < 18 ? 'Good afternoon' : 'Good evening';
	const name = currentProfile().name;
	return name && name !== 'Your profile' ? `${part}, ${name.split(/\s+/u)[0]}` : part;
}

export function mountChatPage({ host, deps, onRoute }) {
	host.innerHTML = `<aside class="chat-history" aria-label="Chat history">
	<div class="chat-history-top">
		<button type="button" class="primary chat-new" data-chat-new>${icon('add')}<span>New chat</span></button>
		<label class="chat-search">${icon('search')}<input type="search" placeholder="Search chats" aria-label="Search chats" /></label>
	</div>
	<nav class="chat-list"></nav>
</aside>
<div class="chat-main">
	<header class="chat-header">
		<button type="button" class="chat-history-toggle" aria-label="Show chat history" aria-expanded="false">${icon('history')}</button>
		<h2 class="chat-title">New chat</h2>
	</header>
	<div class="chat-page-chat"></div>
</div>`;
	const list = host.querySelector('.chat-list');
	const search = host.querySelector('.chat-search input');
	const title = host.querySelector('.chat-title');
	const toggle = host.querySelector('.chat-history-toggle');
	const hero = document.createElement('div');
	hero.className = 'chat-hero';
	const chat = createChat({
		root: host.querySelector('.chat-page-chat'),
		prefix: 'chat',
		deps,
		empty: hero,
		placeholder: 'Ask anything, or ask for a document, workbook or deck',
		onChange: (conversation) => {
			title.textContent = conversation?.title ?? 'New chat';
			onRoute(conversation?.id);
			void renderList();
		},
	});
	function renderHero() {
		hero.innerHTML = `<svg class="rings" viewBox="0 0 100 100" aria-hidden="true"><circle cx="50" cy="50" r="47"/><circle cx="50" cy="50" r="33"/><circle cx="50" cy="50" r="19"/></svg>
<span class="chat-hero-mark" aria-hidden="true">${icon('spark')}</span>
<h1>${escape(greeting())}</h1>
<p>What can I help you create today? I can draft documents, build workbooks and decks, and work across the files in your library.</p>
<div class="chat-starters">${starters
			.map(
				(s, i) =>
					`<button type="button" data-starter="${i}">${s.kind === 'files' ? `<span class="app-badge chat-files-badge">${icon('clip')}</span>` : badge(s.kind)}<span><strong>${escape(s.title)}</strong><small>${escape(s.prompt)}</small></span></button>`,
			)
			.join('')}</div>`;
	}
	renderHero();
	hero.addEventListener('click', (event) => {
		const starter = starters[event.target.closest('[data-starter]')?.dataset.starter];
		if (!starter) return;
		if (starter.kind === 'files' && !chat.hasAttachments()) {
			$('chat-context').click();
			$('chat-input').value = starter.prompt;
			return;
		}
		chat.send(starter.prompt);
	});

	async function renderList() {
		const query = search.value.trim().toLowerCase();
		const all = (await conversations.list()).filter(
			(c) =>
				!query ||
				c.title.toLowerCase().includes(query) ||
				c.messages.some((m) => m.content?.toLowerCase().includes(query)),
		);
		list.innerHTML = all.length
			? groupByDate(all)
					.map(
						(group) =>
							`<span class="nav-caption">${group.label}</span>` +
							group.items
								.map(
									(c) =>
										`<div class="chat-list-item" aria-current="${c.id === chat.id}"><button type="button" data-chat-open="${c.id}" title="${escape(c.title)}">${escape(c.title)}</button><span class="chat-item-actions"><button type="button" data-chat-pin="${c.id}" aria-label="${c.pinned ? 'Unpin' : 'Pin'} ${escape(c.title)}" title="${c.pinned ? 'Unpin' : 'Pin'}">${icon('pin')}</button><button type="button" data-chat-rename="${c.id}" aria-label="Rename ${escape(c.title)}" title="Rename">${icon('edit')}</button><button type="button" data-chat-delete="${c.id}" aria-label="Delete ${escape(c.title)}" title="Delete">${icon('trash')}</button></span></div>`,
								)
								.join(''),
					)
					.join('')
			: `<p class="chat-list-empty">${query ? 'No chats match your search.' : 'Your conversations appear here.'}</p>`;
	}
	search.oninput = () => void renderList();
	conversations.onChange(() => void renderList());
	const setHistory = (open) => {
		host.classList.toggle('history-open', open);
		toggle.setAttribute('aria-expanded', String(open));
	};
	toggle.onclick = () => setHistory(!host.classList.contains('history-open'));
	host.addEventListener('click', (event) => {
		const b = event.target.closest('button');
		if (!b) return;
		if (b.hasAttribute('data-chat-new')) {
			chat.reset();
			setHistory(false);
			chat.focus();
		}
		if (b.dataset.chatOpen) {
			setHistory(false);
			task(() => chat.load(b.dataset.chatOpen));
		}
		if (b.dataset.chatPin)
			task(async () => {
				const c = await conversations.get(b.dataset.chatPin);
				await conversations.put({ ...c, pinned: !c.pinned });
			});
		if (b.dataset.chatRename)
			task(async () => {
				const c = await conversations.get(b.dataset.chatRename);
				const form = document.createElement('form');
				form.innerHTML = `<label>Name<input name="title" value="${escape(c.title)}" maxlength="120" required /></label><button class="primary">Rename</button>`;
				form.onsubmit = (e) => {
					e.preventDefault();
					task(async () => {
						const next = String(new FormData(form).get('title')).trim();
						await conversations.put({ ...c, title: next || c.title });
						if (chat.id === c.id) title.textContent = next || c.title;
						$('dialog').close();
					});
				};
				choose('Rename chat', form);
			});
		if (b.dataset.chatDelete)
			task(async () => {
				const c = await conversations.get(b.dataset.chatDelete);
				const form = document.createElement('form');
				form.innerHTML = `<p>Delete "${escape(c.title)}"? Files it created stay in your library.</p><button class="primary">Delete chat</button>`;
				form.onsubmit = (e) => {
					e.preventDefault();
					task(async () => {
						await conversations.remove(c.id);
						if (chat.id === c.id) chat.reset();
						$('dialog').close();
					});
				};
				choose('Delete chat', form);
			});
	});
	chat.render();
	void renderList();
	// A workspace whose policy keeps no history shows no history list.
	void loadPolicy().then(() => {
		host.querySelector('.chat-history').hidden = !feature('history');
	});
	return {
		show(id) {
			renderHero();
			if (id && id !== chat.id) return chat.load(id);
			if (!chat.id) chat.render();
		},
		ask(prompt, ids = []) {
			chat.reset();
			void chat.attach(ids);
			chat.send(prompt);
		},
		refreshList: renderList,
		get id() {
			return chat.id;
		},
	};
}
