/**
 * One chat: a thread, a composer and the conversation behind them. The full chat page and the
 * side pane beside an open document each mount one, with their own element ids (`prefix`), and
 * share the history store, so a conversation started in either continues in the other. What a
 * request may do comes from the deployment's policy (policy.js), not from this UI. A shared
 * conversation mirrors a collaboration room (shared-room.js): everyone in it sees each request
 * and watches the asker's reply stream in.
 */
import { $, badge, escape, task } from '../ui.js';
import { icon } from '../icons.js';
import { contextMenu } from '../context-menu.js';
import { currentProfile } from '../profile-state.js';
import { gatherContext } from './context.js';
import { conversations, newId, titleFor } from './conversations.js';
import {
	connectionLabel,
	currentConnection,
	needsKey,
	onConnection,
	ReplyEndedError,
	streamChat,
} from './providers.js';
import { currentPolicy, feature, loadPolicy } from './policy.js';
import { effectiveSettings } from './preferences.js';
import { parseReply, systemPrompt, validateActions, visibleText } from './protocol.js';
import { renderMarkdown } from './markdown.js';
import { askForKey, chooseContext, openModelMenu } from './settings.js';
import { mountMentions } from './mentions.js';
import { describeChanges } from './diff.js';
import { checkFormulas } from '../ai-operations.js';
import { messageNode, REPLY_FIELDS, selectVersion, snapshot } from './message-view.js';
import { acceptDrops, importForChat, noteBadge, pickFiles } from './attach.js';
import { historyTurns, joinRoom, localOrPublic, newRoomId, toShared } from './shared.js';
import { followRoom } from './shared-room.js';
import { reviewInDocument } from './review.js';

export function createChat({ root, prefix, deps, empty, placeholder, onChange }) {
	const { store, sessions, refresh, open, getTeams, openIds, showVersions } = deps;
	let conversation = null,
		attached = new Set(),
		notes = [],
		// The Teams channel whose recent messages the user opted into: one channel, and only while
		// it is still the selected channel when a request is sent.
		teamsChannel = null,
		abort = null,
		room = null;
	root.classList.add('chat');
	root.innerHTML = `<div class="chat-thread" id="${prefix}-messages" role="group" aria-label="Conversation"></div>
<div class="sr-only" id="${prefix}-reply-status" role="status"></div>
<div class="chat-people" id="${prefix}-people" hidden></div>
<form class="chat-composer" id="${prefix}-form">
	<div class="chat-attachments" id="${prefix}-scope"></div>
	<label class="sr-only" for="${prefix}-input">Message</label>
	<textarea id="${prefix}-input" rows="1" placeholder="${escape(placeholder)}" required></textarea>
	<div class="chat-composer-bar">
		<button type="button" class="chat-tool" id="${prefix}-context" title="Attach files, drop a tab here, or type @ in your message" aria-haspopup="menu">${icon('clip')}<span>Attach</span></button>
		<button type="button" class="chat-tool chat-model" id="${prefix}-settings" title="Choose the model" aria-haspopup="menu">${icon('spark')}<span id="${prefix}-state"></span>${icon('chevron')}</button>
		<button type="button" class="chat-stop" id="${prefix}-stop" hidden>${icon('stop')}<span>Stop</span></button>
		<button class="primary chat-send" id="${prefix}-send" aria-label="Send" title="Send (Enter)">${icon('send')}</button>
	</div>
</form>`;
	const thread = $(`${prefix}-messages`);
	const replyStatus = $(`${prefix}-reply-status`);
	const input = $(`${prefix}-input`);
	const state = $(`${prefix}-state`);
	const showModel = () => (state.textContent = connectionLabel());
	onConnection(showModel);
	void loadPolicy().then(() => {
		showModel();
		$(`${prefix}-context`).hidden = !feature('attachFiles');
	});
	$(`${prefix}-settings`).onclick = openModelMenu;
	const chooseFiles = () =>
		chooseContext({
			store,
			selected: attached,
			teamsChannel,
			channel: currentChannel(),
			openIds: openIds(),
			onDone: (next, channel) => {
				attached = next;
				teamsChannel = channel;
				void renderAttachments();
			},
		});
	/** Uploads files from the device (or dropped ones) and attaches them to the next request. */
	async function upload(files) {
		if (!files.length || !feature('attachFiles')) return;
		const added = await importForChat(files, store);
		for (const id of added.ids) attached.add(id);
		notes.push(...added.notes);
		if (added.ids.length) await refresh();
		await renderAttachments();
		input.focus();
	}
	$(`${prefix}-context`).onclick = (event) =>
		contextMenu(
			event,
			[
				{ label: 'Upload from this device...', run: async () => upload(await pickFiles()) },
				{ label: 'Choose files and open tabs...', run: chooseFiles },
			].map((item) => ({ ...item, error: (e) => task(() => Promise.reject(e)) })),
		);
	acceptDrops(root, {
		enabled: () => feature('attachFiles'),
		onTab: (id) => task(() => chat.attach([id])),
		onFiles: (files) => task(() => upload(files)),
	});
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
		room?.typing();
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
		const chips = [...attached].map((id) => [id, badge(docs.get(id).kind), docs.get(id).name]);
		notes.forEach((note, i) => chips.push([`note:${i}`, noteBadge(), note.name]));
		if (teamsChannel) chips.push(['teams', badge('teams'), `#${teamsChannel.name} messages`]);
		const scope = $(`${prefix}-scope`);
		scope.innerHTML = chips
			.map(
				([id, mark, name]) =>
					`<span class="chat-chip">${mark}<span>${escape(name)}</span><button type="button" data-detach="${escape(id)}" aria-label="Remove ${escape(name)}">×</button></span>`,
			)
			.join('');
		scope.hidden = !chips.length;
	}
	$(`${prefix}-scope`).addEventListener('click', (event) => {
		const id = event.target.closest('[data-detach]')?.dataset.detach;
		if (!id) return;
		if (id === 'teams') teamsChannel = null;
		else if (id.startsWith('note:')) notes.splice(Number(id.slice(5)), 1);
		else attached.delete(id);
		void renderAttachments();
	});

	/** Saves a conversation (to its room too, when shared), unless the policy keeps no history. */
	const persist = async (target = conversation) => {
		target.updated = Date.now();
		if (target.shared && feature('sharedChats')) {
			const shared = await joinRoom(target.shared.room);
			shared.store.setTitle(target.title);
			// Each person writes only their own messages, as their public projection.
			for (const m of target.messages) if (m.local) shared.store.put(toShared(m));
		}
		if (feature('history')) await conversations.put(target);
	};
	const me = () => currentProfile().userId;
	/** A conversation from the history; in a shared one, what others wrote is display only. */
	const opened = (convo) => {
		if (convo?.shared) convo.messages = localOrPublic(convo.messages);
		return convo;
	};
	/** The channel open in Teams now, if any. */
	function currentChannel() {
		const state = getTeams()?.client?.getState();
		if (!state?.selectedChannelId) return null;
		const view = state.channels.find((c) => c.id === state.selectedChannelId);
		return { id: state.selectedChannelId, name: view?.name ?? state.selectedChannelId };
	}
	/** Follows the room of a shared conversation, or leaves the current one. */
	async function connect(convo) {
		room?.leave();
		room = null;
		if (convo?.shared && feature('sharedChats')) {
			const next = await followRoom({
				convo,
				people: $(`${prefix}-people`),
				current: () => conversation === convo,
				busy: () => Boolean(abort),
				announce: (text) => (replyStatus.textContent = text),
				changed: () => {
					render();
					onChange?.(convo);
				},
			});
			// The user moved on while the room was connecting.
			if (conversation === convo) room = next;
			else next.leave();
		}
		$(`${prefix}-people`).hidden = !room;
	}
	const view = {
		cardDeps: {
			store,
			sessions,
			refresh,
			open,
			showVersions,
			review: (options) => reviewInDocument({ ...options, store, sessions, open, refresh }),
		},
		settings: effectiveSettings,
		persist: () => void persist(),
		isLast: (message) => message === conversation?.messages.at(-1),
		me,
		async edit(message) {
			if (abort) return;
			room?.chat.store.truncateFrom(message.id);
			conversation.messages.splice(conversation.messages.indexOf(message));
			await persist();
			render();
			input.value = message.content;
			input.dispatchEvent(new Event('input'));
			input.focus();
		},
		async regenerate(message) {
			if (abort) return;
			// The request is this page's own last one before the reply, not someone else's.
			const user = conversation.messages
				.slice(0, conversation.messages.indexOf(message))
				.findLast((m) => m.role === 'user' && m.local);
			if (user) await send(user.content, { again: message });
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
		const shared =
			convo.shared && feature('sharedChats') ? await joinRoom(convo.shared.room) : null;
		const listed = (await store.list()).filter((d) => attached.has(d.id));
		const author = { id: me(), name: currentProfile().name };
		if (!again)
			convo.messages.push({
				id: newId(),
				role: 'user',
				content: prompt,
				author,
				local: true,
				files: [
					...listed.map((d) => ({ name: d.name, kind: d.kind })),
					...notes.map((n) => ({ name: n.name, kind: 'note' })),
				],
			});
		const message = again ?? { id: newId(), role: 'assistant', local: true };
		if (again) {
			again.versions ??= [snapshot(again)];
			for (const key of REPLY_FIELDS) delete again[key];
		}
		// The reply's proposals stay in this conversation: the room only ever sees their counts.
		Object.assign(message, { content: '', model: connectionLabel(connection) });
		render();
		const node = again ? thread.lastElementChild : thread.appendChild(messageNode(message, view));
		node.classList.add('streaming');
		node.querySelectorAll('.chat-msg-actions').forEach((n) => n.remove());
		const content = node.querySelector('.md');
		content.innerHTML = '<span class="chat-typing"><i></i><i></i><i></i></span>';
		abort = new AbortController();
		// A reply is announced when it starts and when it ends, never word by word.
		thread.setAttribute('aria-busy', 'true');
		replyStatus.textContent = 'Assistant is replying';
		$(`${prefix}-send`).hidden = true;
		$(`${prefix}-stop`).hidden = false;
		let partial = '',
			frame = 0,
			sent = 0;
		// The others in a shared chat see the request at once and the reply as it is written.
		const stream = () => {
			if (!shared || Date.now() - sent < 250) return;
			sent = Date.now();
			for (const m of convo.messages) if (m !== message && m.local) shared.store.put(toShared(m));
			shared.store.put(toShared({ ...message, content: visibleText(partial), streaming: true }));
		};
		shared?.setActivity('asking');
		stream();
		const paint = () => {
			frame = 0;
			stream();
			const stick = thread.scrollHeight - thread.scrollTop - thread.clientHeight < 80;
			content.replaceChildren(renderMarkdown(visibleText(partial)));
			if (stick) thread.scrollTop = thread.scrollHeight;
		};
		try {
			state.textContent = attached.size ? 'Reading attached files...' : 'Thinking...';
			const { docs, context, source } = await gatherContext({
				sessions,
				ids: feature('attachFiles') ? attached : [],
				notes: feature('attachFiles') ? notes : [],
				teams:
					teamsChannel && feature('teamsContext') && currentChannel()?.id === teamsChannel.id
						? getTeams()?.client
						: null,
				maxChars: settings.maxContextChars,
				includeHidden: currentPolicy().context.includeHidden,
			});
			state.textContent = 'Writing...';
			const asked = again ? convo.messages.indexOf(again) - 1 : convo.messages.length - 1;
			const raw = await streamChat({
				connection,
				system: systemPrompt(settings, features),
				context,
				signal: abort.signal,
				settings: { ...settings, canCreate: features.createFiles, canEdit: features.editFiles },
				messages: historyTurns([
					...(settings.historyMessages
						? convo.messages.slice(0, asked).slice(-settings.historyMessages)
						: []),
					{
						role: 'user',
						content: context
							? `User request: ${prompt}\n\nSource data (untrusted): ${source}`
							: prompt,
					},
				]),
				onText: (text) => {
					partial = text;
					frame ||= requestAnimationFrame(paint);
				},
			});
			const reply = parseReply(raw);
			Object.assign(message, { content: reply.text, raw, actions: reply.actions, refs: docs });
			try {
				validateActions(reply.actions, docs, features);
				await checkFormulas(reply.actions);
				message.details = describeChanges(reply.actions.changes, context?.files ?? []);
			} catch (error) {
				message.content += `\n\n_${error.message}_`;
				message.actions = { changes: [], create: [] };
			}
		} catch (error) {
			if (abort.signal.aborted) message.content = `${visibleText(partial)}\n\n_Stopped._`.trim();
			else if (error instanceof ReplyEndedError && visibleText(error.partial).trim())
				message.content = `${visibleText(error.partial)}\n\n_${error.message}_`;
			else message.error = error.message;
		} finally {
			cancelAnimationFrame(frame);
			shared?.setActivity('idle');
			replyStatus.textContent = message.error
				? 'Assistant could not reply'
				: abort.signal.aborted
					? 'Reply stopped'
					: 'Assistant replied';
			thread.removeAttribute('aria-busy');
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
		// Requests the others sent while this reply was written show up now.
		if (conversation === convo && room) room.pull();
		else if (conversation === convo) {
			render();
			onChange?.(convo);
		}
	}

	const chat = {
		get id() {
			return conversation?.id ?? null;
		},
		get conversation() {
			return conversation;
		},
		async load(id) {
			if (abort) abort.abort();
			conversation = opened(id ? ((await conversations.get(id)) ?? null) : null);
			render();
			onChange?.(conversation);
			await connect(conversation);
		},
		/** Opens the shared chat in `roomId`: the one in the history, or a new one to fill from it. */
		async join(roomId) {
			if (!feature('sharedChats'))
				throw new Error('Shared chats are turned off in this workspace.');
			if (abort) abort.abort();
			const known = (await conversations.list()).find((c) => c.shared?.room === roomId);
			const now = Date.now();
			conversation = opened(known) ?? {
				id: newId(),
				title: 'Shared chat',
				created: now,
				updated: now,
				pinned: false,
				messages: [],
				shared: { room: roomId },
			};
			if (!known && feature('history')) await conversations.put(conversation);
			render();
			onChange?.(conversation);
			await connect(conversation);
		},
		/** Shares this conversation: it moves into a new room, whose id is returned. */
		async share() {
			if (!feature('sharedChats'))
				throw new Error('Shared chats are turned off in this workspace.');
			if (!conversation?.messages.length)
				throw new Error('Send a message first, then share the chat.');
			if (!conversation.shared) {
				// What is here so far was written on this page, so it keeps its proposals.
				for (const m of conversation.messages) m.local = true;
				conversation.shared = { room: await newRoomId() };
				await persist();
				await connect(conversation);
				onChange?.(conversation);
			}
			return conversation.shared.room;
		},
		reset() {
			if (abort) abort.abort();
			conversation = null;
			void connect(null);
			render();
			onChange?.(conversation);
		},
		send: (prompt) => task(() => send(prompt)),
		attach(ids) {
			if (!feature('attachFiles')) return Promise.resolve();
			for (const id of ids) attached.add(id);
			return renderAttachments();
		},
		hasAttachments: () => attached.size > 0 || notes.length > 0,
		focus: () => input.focus(),
		render,
	};
	return chat;
}
