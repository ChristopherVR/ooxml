/**
 * One chat: a thread, a composer and the conversation behind them. The full chat page and the
 * side pane beside an open document each mount one, with their own element ids (`prefix`), and
 * share the history store, so a conversation started in either continues in the other. What a
 * request may do comes from the deployment's policy (policy.js), not from this UI.
 */
import { $, badge, escape, task } from '../ui.js';
import { icon } from '../icons.js';
import { gatherContext } from './context.js';
import { conversations, newId, titleFor } from './conversations.js';
import {
	connectionLabel,
	currentConnection,
	needsKey,
	onConnection,
	streamChat,
} from './providers.js';
import { currentPolicy, feature, loadPolicy } from './policy.js';
import { effectiveSettings } from './preferences.js';
import { parseReply, systemPrompt, validateActions, visibleText } from './protocol.js';
import { renderMarkdown } from './markdown.js';
import { askForKey, chooseContext, openModelMenu } from './settings.js';
import { mountMentions } from './mentions.js';
import { describeChanges } from './diff.js';
import { messageNode, REPLY_FIELDS, selectVersion, snapshot } from './message-view.js';

export function createChat({ root, prefix, deps, empty, placeholder, onChange }) {
	const { store, sessions, refresh, open, getTeams, openIds, showVersions } = deps;
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
		<button type="button" class="chat-tool" id="${prefix}-context" title="Attach files, or type @ in your message">${icon('clip')}<span>Attach</span></button>
		<button type="button" class="chat-tool chat-model" id="${prefix}-settings" title="Choose the model" aria-haspopup="menu">${icon('spark')}<span id="${prefix}-state"></span>${icon('chevron')}</button>
		<button type="button" class="chat-stop" id="${prefix}-stop" hidden>${icon('stop')}<span>Stop</span></button>
		<button class="primary chat-send" id="${prefix}-send" aria-label="Send" title="Send (Enter)">${icon('send')}</button>
	</div>
</form>`;
	const thread = $(`${prefix}-messages`);
	const input = $(`${prefix}-input`);
	const state = $(`${prefix}-state`);
	const showModel = () => (state.textContent = connectionLabel());
	onConnection(showModel);
	void loadPolicy().then(() => {
		showModel();
		$(`${prefix}-context`).hidden = !feature('attachFiles');
	});
	$(`${prefix}-settings`).onclick = openModelMenu;
	$(`${prefix}-context`).onclick = () =>
		task(() =>
			chooseContext({
				store,
				selected: attached,
				includeChat,
				openIds: openIds(),
				onDone: (next, chat) => {
					attached = next;
					includeChat = chat;
					void renderAttachments();
				},
			}),
		);
	mountMentions({
		input,
		items: async () => {
			if (!feature('attachFiles')) return [];
			const docs = (await store.list()).filter((d) => !d.parent);
			const ids = openIds();
			return [
				...docs.filter((d) => ids.includes(d.id)).map((d) => ({ ...d, group: 'Open tabs' })),
				...docs.filter((d) => !ids.includes(d.id)).map((d) => ({ ...d, group: 'Library' })),
			];
		},
		onPick: (doc) => {
			attached.add(doc.id);
			void renderAttachments();
		},
	});
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

	/** Saves a conversation, unless the policy keeps chats out of history. */
	const persist = async (target = conversation) => {
		target.updated = Date.now();
		if (feature('history')) await conversations.put(target);
	};
	const view = {
		cardDeps: { store, sessions, refresh, open, showVersions },
		settings: effectiveSettings,
		persist: () => void persist(),
		isLast: (message) => message === conversation?.messages.at(-1),
		async edit(message) {
			if (abort) return;
			conversation.messages.splice(conversation.messages.indexOf(message));
			await persist();
			render();
			input.value = message.content;
			input.dispatchEvent(new Event('input'));
			input.focus();
		},
		async regenerate(message) {
			if (abort) return;
			const user = conversation.messages.at(-2);
			await send(user.content, { again: message });
		},
		async switchVersion(message, index) {
			selectVersion(message, index);
			await persist();
			render();
		},
	};
	function render() {
		thread.replaceChildren(...(conversation?.messages ?? []).map((m) => messageNode(m, view)));
		root.toggleAttribute('data-empty', !conversation?.messages.length);
		if (!conversation?.messages.length && empty) thread.append(empty);
		thread.scrollTop = thread.scrollHeight;
	}

	/** Asks the model; `again` is an earlier reply to regenerate as a new version of itself. */
	async function send(prompt, { again } = {}) {
		if (abort) return;
		await loadPolicy();
		const connection = currentConnection();
		if (!connection)
			throw new Error('No model is available in this workspace. Ask your administrator.');
		if (needsKey(connection)) {
			askForKey(connection);
			throw new Error(`Enter your key for ${connection.label} to continue.`);
		}
		const settings = effectiveSettings();
		const features = currentPolicy().features;
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
		const message = again ?? { id: newId(), role: 'assistant' };
		if (again) {
			again.versions ??= [snapshot(again)];
			for (const key of REPLY_FIELDS) delete again[key];
		}
		Object.assign(message, { content: '', model: connectionLabel(connection) });
		render();
		const node = again ? thread.lastElementChild : thread.appendChild(messageNode(message, view));
		node.classList.add('streaming');
		node.querySelectorAll('.chat-msg-actions').forEach((n) => n.remove());
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
				ids: feature('attachFiles') ? attached : [],
				teams: includeChat && feature('teamsContext') ? getTeams()?.client : null,
				maxChars: settings.maxContextChars,
			});
			state.textContent = 'Writing...';
			const asked = again ? convo.messages.indexOf(again) - 1 : convo.messages.length - 1;
			const history = [];
			for (const m of convo.messages.slice(0, asked).slice(-settings.historyMessages)) {
				if (m.error) history.pop();
				else history.push({ role: m.role, content: m.raw ?? m.content });
			}
			const raw = await streamChat({
				connection,
				system: systemPrompt(settings, features),
				context,
				signal: abort.signal,
				settings: { ...settings, canCreate: features.createFiles, canEdit: features.editFiles },
				messages: [
					...(settings.historyMessages ? history : []),
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
				validateActions(reply.actions, docs, features);
				message.details = describeChanges(reply.actions.changes, context?.files ?? []);
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
		if (again) {
			again.versions.push(snapshot(again));
			again.selected = again.versions.length - 1;
		} else convo.messages.push(message);
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
			if (!feature('attachFiles')) return Promise.resolve();
			for (const id of ids) attached.add(id);
			return renderAttachments();
		},
		hasAttachments: () => attached.size > 0,
		focus: () => input.focus(),
		render,
	};
}
