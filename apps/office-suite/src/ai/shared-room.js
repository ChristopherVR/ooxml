/**
 * Following a shared chat's room from one chat view: its messages and title become the
 * conversation's, the history copy is refreshed, and the bar above the composer shows who else is
 * here and what they are doing.
 */
import { escape } from '../ui.js';
import { icon } from '../icons.js';
import { conversations } from './conversations.js';
import { feature } from './policy.js';
import { fromShared, joinRoom } from './shared.js';

export async function followRoom({ convo, people, current, busy, changed }) {
	const chat = await joinRoom(convo.shared.room);
	let saving = 0,
		typing = 0;
	const pull = () => {
		// A reply being written here owns the message list until it ends, then pulls again.
		if (busy()) return;
		const rows = chat.store.messages();
		// An empty room that has not synced yet is not an empty chat.
		if (!rows.length && !chat.session.synced) return;
		convo.messages = rows.map(fromShared);
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
			active
				? `${escape(active.name)} is ${active.activity === 'asking' ? 'asking the assistant' : 'typing'}...`
				: others.length
					? `${others.length} other${others.length === 1 ? '' : 's'} here`
					: 'Waiting for others to join'
		}</span>`;
	};
	const stopRows = chat.store.observe(pull);
	const stopPeople = chat.onPeople(paint);
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
			clearTimeout(saving);
			clearTimeout(typing);
			chat.setActivity('idle');
		},
	};
}
