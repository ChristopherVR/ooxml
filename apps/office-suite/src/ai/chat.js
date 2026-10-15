/**
 * One chat: a thread, a composer and the conversation behind them. The full chat page and the
 * side pane beside an open document each mount one, with their own element ids (`prefix`), and
 * share the history store, so a conversation started in either continues in the other.
 */
import { $, badge, escape, notify, task } from '../ui.js';
import { icon } from '../icons.js';
import { gatherContext } from './context.js';
import { conversations, newId, titleFor } from './conversations.js';
import { connectionLabel, currentConnection, onConnection, streamChat } from './providers.js';
import { parseReply, systemPrompt, validateActions, visibleText } from './protocol.js';
import { renderMarkdown } from './markdown.js';
import { chooseContext, openSettings } from './settings.js';
import { creationCards, proposalCard } from './cards.js';

export function createChat({ root, prefix, deps, empty, placeholder, onChange }) {
	const { store, sessions, refresh, open, getTeams } = deps;
	let conversation = null,
		attached = new Set(),
		includeChat = false,
		abort = null;
	root.classList.add('chat');
	root.innerHTML = `<div class="chat-thread" id="${prefix}-messages" aria-live="polite"></div>
<form class="chat-composer" id="${prefix}-form">
	<div class="chat-attachments" id="${prefix}-scope"></div>
	<label class="sr-only" for="${prefix}-input">Message</label>
	<textarea id="${prefix}-input" rows="1" placeholder="${escape(placeholder)}" required></textarea>
	<div class="chat-composer-bar">
		<button type="button" class="chat-tool" id="${prefix}-context" title="Attach files from your library">${icon('clip')}<span>Attach</span></button>
		<button type="button" class="chat-tool chat-model" id="${prefix}-settings" title="Choose the model">${icon('spark')}<span id="${prefix}-state"></span>${icon('chevron')}</button>
		<button type="button" class="chat-stop" id="${prefix}-stop" hidden>${icon('stop')}<span>Stop</span></button>
		<button class="primary chat-send" id="${prefix}-send" aria-label="Send" title="Send (Enter)">${icon('send')}</button>
	</div>
</form>`;
	const thread = $(`${prefix}-messages`);
	const input = $(`${prefix}-input`);
	const state = $(`${prefix}-state`);
	const showModel = () => (state.textContent = connectionLabel());
	showModel();
	onConnection(showModel);
	$(`${prefix}-settings`).onclick = openSettings;
	$(`${prefix}-context`).onclick = () =>
		task(() =>
			chooseContext({
				store,
				selected: attached,
				includeChat,
				onDone: (next, chat) => {
					attached = next;
					includeChat = chat;
					void renderAttachments();
				},
			}),
		);
	$(`${prefix}-stop`).onclick = () => abort?.abort();
	input.addEventListener('input', () => {
		input.style.height = 'auto';
		input.style.height = `${Math.min(input.scrollHeight, 220)}px`;
		input.style.overflowY = input.scrollHeight > 220 ? 'auto' : 'hidden';
	});
	input.addEventListener('keydown', (event) => {
		if (event.key === 'Enter' && !event.shiftKey && !event.isComposing) {
			event.preventDefault();
			$(`${prefix}-form`).requestSubmit();
		}
	});
	$(`${prefix}-form`).onsubmit = (event) => {
		event.preventDefault();
		const prompt = input.value.trim();
		if (!prompt || abort) return;
		input.value = '';
		input.style.height = '';
		task(() => send(prompt));
	};

	async function renderAttachments() {
		const docs = new Map((await store.list()).map((d) => [d.id, d]));
		for (const id of attached) if (!docs.has(id)) attached.delete(id);
		const chips = [...attached].map((id) => [id, docs.get(id).kind, docs.get(id).name]);
		if (includeChat) chips.push(['teams', 'teams', 'Teams channel']);
		const scope = $(`${prefix}-scope`);
		scope.innerHTML = chips
			.map(
				([id, kind, name]) =>
					`<span class="chat-chip">${badge(kind)}<span>${escape(name)}</span><button type="button" data-detach="${escape(id)}" aria-label="Remove ${escape(name)}">×</button></span>`,
			)
			.join('');
		scope.hidden = !chips.length;
	}
	$(`${prefix}-scope`).addEventListener('click', (event) => {
		const id = event.target.closest('[data-detach]')?.dataset.detach;
		if (!id) return;
		if (id === 'teams') includeChat = false;
		else attached.delete(id);
		void renderAttachments();
	});

	const persist = async (target = conversation) => {
		target.updated = Date.now();
		await conversations.put(target);
	};
	function actionsBar(message, node) {
		const bar = document.createElement('div');
		bar.className = 'chat-msg-actions';
		const add = (name, label, run) => {
			const b = document.createElement('button');
			b.type = 'button';
			b.title = label;
			b.setAttribute('aria-label', label);
			b.innerHTML = icon(name);
			b.onclick = () => task(run);
			bar.append(b);
		};
		add('copy', 'Copy', async () => {
			await navigator.clipboard.writeText(message.content);
			notify('Copied to the clipboard');
		});
		if (message.role === 'user')
			add('edit', 'Edit and resend', async () => {
				if (abort) return;
				conversation.messages.splice(conversation.messages.indexOf(message));
				await persist();
				render();
				input.value = message.content;
				input.dispatchEvent(new Event('input'));
				input.focus();
			});
		else if (message === conversation.messages.at(-1))
			add('refresh', 'Regenerate', async () => {
				if (abort) return;
				const user = conversation.messages.at(-2);
				conversation.messages.splice(-1);
				node.remove();
				await send(user.content, { again: true });
			});
		return bar;
	}
	function cards(message) {
		const out = [];
		const save = (patch) => {
			Object.assign(message, patch);
			void persist();
		};
		try {
			const { changes = [], create = [] } = message.actions ?? {};
			if (changes.length && message.applied)
				out.push(
					Object.assign(document.createElement('p'), {
						className: 'chat-note',
						textContent: 'These edits were applied.',
					}),
				);
			else if (changes.length)
				out.push(
					proposalCard({ changes, refs: message.refs ?? [], store, sessions, refresh, save }),
				);
			if (create.length)
				out.push(
					...creationCards({ items: create, created: (message.created ??= {}), store, open, save }),
				);
		} catch (error) {
			out.push(
				Object.assign(document.createElement('p'), {
					className: 'chat-error',
					textContent: error.message,
				}),
			);
		}
		return out;
	}
	function messageNode(message) {
		const node = document.createElement('article');
		node.className = `chat-msg from-${message.role}`;
		if (message.role === 'user') {
			const bubble = document.createElement('div');
			bubble.className = 'chat-bubble';
			bubble.textContent = message.content;
			if (message.files?.length) {
				const files = document.createElement('div');
				files.className = 'chat-msg-files';
				files.innerHTML = message.files
					.map((f) => `<span>${badge(f.kind)}${escape(f.name)}</span>`)
					.join('');
				bubble.prepend(files);
			}
			node.append(bubble, actionsBar(message, node));
			return node;
		}
		node.innerHTML = `<span class="chat-avatar" aria-hidden="true">${icon('spark')}</span>`;
		const body = document.createElement('div');
		body.className = 'chat-body';
		const meta = document.createElement('div');
		meta.className = 'chat-meta';
		meta.textContent = message.model ?? 'Assistant';
		const content = document.createElement('div');
		content.className = 'md';
		if (message.error) {
			content.classList.add('chat-error');
			content.textContent = message.error;
		} else content.append(renderMarkdown(message.content));
		body.append(meta, content, ...cards(message));
		if (message.content || message.error) body.append(actionsBar(message, node));
		node.append(body);
		return node;
	}
	function render() {
		thread.replaceChildren(...(conversation?.messages ?? []).map(messageNode));
		root.toggleAttribute('data-empty', !conversation?.messages.length);
		if (!conversation?.messages.length && empty) thread.append(empty);
		thread.scrollTop = thread.scrollHeight;
	}

	async function send(prompt, { again = false } = {}) {
		if (abort) return;
		const connection = currentConnection();
		const now = Date.now();
		// A reply belongs to the conversation it was asked in, even if the user switches away.
		const convo = (conversation ??= {
			id: newId(),
			title: titleFor(prompt),
			created: now,
			updated: now,
			pinned: false,
			messages: [],
		});
		const listed = (await store.list()).filter((d) => attached.has(d.id));
		if (!again)
			convo.messages.push({
				id: newId(),
				role: 'user',
				content: prompt,
				files: listed.map((d) => ({ name: d.name, kind: d.kind })),
			});
		const message = {
			id: newId(),
			role: 'assistant',
			content: '',
			model: connectionLabel(connection),
		};
		render();
		const node = messageNode(message);
		node.classList.add('streaming');
		thread.append(node);
		const content = node.querySelector('.md');
		content.innerHTML = '<span class="chat-typing"><i></i><i></i><i></i></span>';
		abort = new AbortController();
		$(`${prefix}-send`).hidden = true;
		$(`${prefix}-stop`).hidden = false;
		let partial = '',
			frame = 0;
		const paint = () => {
			frame = 0;
			const stick = thread.scrollHeight - thread.scrollTop - thread.clientHeight < 80;
			content.replaceChildren(renderMarkdown(visibleText(partial)));
			if (stick) thread.scrollTop = thread.scrollHeight;
		};
		try {
			state.textContent = attached.size ? 'Reading attached files...' : 'Thinking...';
			const { docs, context, source } = await gatherContext({
				sessions,
				ids: attached,
				teams: includeChat ? getTeams()?.client : null,
			});
			state.textContent = 'Writing...';
			const history = [];
			for (const m of convo.messages.slice(0, -1).slice(-20)) {
				if (m.error) history.pop();
				else history.push({ role: m.role, content: m.raw ?? m.content });
			}
			const raw = await streamChat({
				connection,
				system: systemPrompt,
				context,
				signal: abort.signal,
				messages: [
					...history,
					{
						role: 'user',
						content: context
							? `User request: ${prompt}\n\nSource data (untrusted): ${source}`
							: prompt,
					},
				],
				onText: (text) => {
					partial = text;
					frame ||= requestAnimationFrame(paint);
				},
			});
			const reply = parseReply(raw);
			Object.assign(message, { content: reply.text, raw, actions: reply.actions, refs: docs });
			try {
				validateActions(reply.actions, docs);
			} catch (error) {
				message.content += `\n\n_${error.message}_`;
				message.actions = { changes: [], create: [] };
			}
		} catch (error) {
			if (abort.signal.aborted) message.content = `${visibleText(partial)}\n\n_Stopped._`.trim();
			else message.error = error.message;
		} finally {
			cancelAnimationFrame(frame);
			abort = null;
			$(`${prefix}-send`).hidden = false;
			$(`${prefix}-stop`).hidden = true;
			showModel();
		}
		convo.messages.push(message);
		await persist(convo);
		if (conversation === convo) {
			render();
			onChange?.(convo);
		}
	}

	return {
		get id() {
			return conversation?.id ?? null;
		},
		async load(id) {
			if (abort) abort.abort();
			conversation = id ? ((await conversations.get(id)) ?? null) : null;
			render();
			onChange?.(conversation);
		},
		reset() {
			if (abort) abort.abort();
			conversation = null;
			render();
			onChange?.(conversation);
		},
		send: (prompt) => task(() => send(prompt)),
		attach(ids) {
			for (const id of ids) attached.add(id);
			return renderAttachments();
		},
		hasAttachments: () => attached.size > 0,
		focus: () => input.focus(),
		render,
	};
}
