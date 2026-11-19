/**
 * Following a shared chat's room from one chat view: its messages and title become the
 * conversation's (this page's own messages keep their local proposals; the others are display
 * only), the history copy is refreshed, and the bar above the composer shows who else is here and
 * what they are doing.
 */
import { escape } from '../ui.js';
import { icon } from '../icons.js';
import { conversations } from './conversations.js';
import { feature } from './policy.js';
import { joinRoom, mergeRoom } from './shared.js';

/**
 * What a screen reader hears when messages arrive from the room: each request someone else sent,
 * and each reply once it is complete (never word by word). `shown` is what this view already had,
 * its own requests and replies included, so only arrivals are read.
 */
export function arrivals(shown, messages) {
	const before = new Map(shown.map((m) => [m.id, Boolean(m.streaming)]));
	const out = [];
	for (const m of messages) {
		const known = before.has(m.id);
		if (m.role === 'user' && !known) out.push(`${m.author?.name || 'Someone'}: ${m.content}`);
		else if (m.role === 'assistant' && !m.streaming && (!known || before.get(m.id)))
			out.push('Assistant replied');
	}
	return out;
}

export async function followRoom({ convo, people, current, busy, changed, announce }) {
	const chat = await joinRoom(convo.shared.room);
	let saving = 0,
		typing = 0,
		primed = false,
		// Why the room is not syncing (a refused credential, an unreachable server), until it connects.
		problem = '';
	const pull = () => {
		// A reply being written here owns the message list until it ends, then pulls again.
		if (busy()) return;
		const rows = chat.store.messages();
		// An empty room that has not synced yet is not an empty chat.
		if (!rows.length && !chat.session.synced) return;
		const messages = mergeRoom(rows, convo.messages);
		// Joining shows the history; only what arrives after that is announced.
		if (primed && current()) for (const text of arrivals(convo.messages, messages)) announce(text);
		primed = true;
		convo.messages = messages;
		convo.title = chat.store.title() || convo.title;
		clearTimeout(saving);
		if (feature('history'))
			saving = setTimeout(() => void conversations.put({ ...convo, updated: Date.now() }), 600);
		if (current()) changed();
	};
	const paint = () => {
		const others = chat.people();
		const active =
			others.find((p) => p.activity === 'asking') ?? others.find((p) => p.activity === 'typing');
		people.innerHTML = `<span class="chat-shared-mark">${icon('people')}Shared chat</span>${others
			.slice(0, 6)
			.map(
				(p) =>
					`<span class="chat-person" style="--person:${escape(p.color)}" title="${escape(p.name)}">${escape(p.name.slice(0, 1).toUpperCase())}</span>`,
			)
			.join('')}<span class="chat-people-state">${
			problem
				? escape(problem)
				: active
					? `${escape(active.name)} is ${active.activity === 'asking' ? 'asking the assistant' : 'typing'}...`
					: others.length
						? `${others.length} other${others.length === 1 ? '' : 's'} here`
						: 'Waiting for others to join'
		}</span>`;
	};
	const stopRows = chat.store.observe(pull);
	const stopPeople = chat.onPeople(paint);
	const stopError = chat.session.on('error', (error) => {
		problem = error.message;
		paint();
	});
	const stopStatus = chat.session.on('status', (status) => {
		if (status !== 'connected' || !problem) return;
		problem = '';
		paint();
	});
	pull();
	paint();
	return {
		chat,
		pull,
		typing() {
			chat.setActivity('typing');
			clearTimeout(typing);
			typing = setTimeout(() => chat.setActivity('idle'), 3000);
		},
		leave() {
			stopRows();
			stopPeople();
			stopError();
			stopStatus();
			clearTimeout(saving);
			clearTimeout(typing);
			chat.setActivity('idle');
		},
	};
}
