/**
 * How one message is drawn: the user's bubble with the files it referenced, or the assistant's
 * reply with its cards, actions and, after regenerating, a switcher between its versions.
 */
import { badge, escape, notify, task } from '../ui.js';
import { icon } from '../icons.js';
import { renderMarkdown } from './markdown.js';
import { creationCards, proposalCard } from './cards.js';
import { noteBadge } from './attach.js';

/** The fields that make up one version of a reply. */
export const REPLY_FIELDS = [
	'content',
	'raw',
	'actions',
	'refs',
	'details',
	'model',
	'error',
	'applied',
	'created',
];
export const snapshot = (message) =>
	Object.fromEntries(
		REPLY_FIELDS.filter((k) => message[k] !== undefined).map((k) => [k, message[k]]),
	);
/** Shows version `index` of a reply, keeping the version on screen up to date first. */
export function selectVersion(message, index) {
	message.versions[message.selected ?? message.versions.length - 1] = snapshot(message);
	for (const key of REPLY_FIELDS) delete message[key];
	Object.assign(message, message.versions[index]);
	message.selected = index;
}

function note(className, text) {
	return Object.assign(document.createElement('p'), { className, textContent: text });
}

const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;

/** What someone else's reply proposed, as counts: its edits and files are theirs alone. */
function remoteNotes({ edits = 0, newFiles = 0, applied = 0 }) {
	const out = [];
	if (edits && applied)
		out.push(
			note(
				'chat-note',
				`${applied} of ${plural(edits, 'proposed edit was', 'proposed edits were')} applied.`,
			),
		);
	else if (edits)
		out.push(
			note(
				'chat-note',
				`${plural(edits, 'edit was', 'edits were')} proposed for the files of the person who asked. Only they can review and apply ${edits === 1 ? 'it' : 'them'}.`,
			),
		);
	if (newFiles)
		out.push(
			note(
				'chat-note',
				`${plural(newFiles, 'new file was', 'new files were')} proposed to the person who asked. Only they can create ${newFiles === 1 ? 'it' : 'them'}.`,
			),
		);
	return out;
}

function cards(message, ctx) {
	if (message.remote) return remoteNotes(message);
	const out = [];
	const save = (patch) => {
		Object.assign(message, patch);
		ctx.persist();
	};
	try {
		const { changes = [], create = [] } = message.actions ?? {};
		if (changes.length && message.applied) {
			const count = Array.isArray(message.applied)
				? message.applied.filter(Boolean).length
				: changes.length;
			out.push(note('chat-note', `${count} of ${changes.length} proposed edits were applied.`));
		} else if (changes.length)
			out.push(
				proposalCard({
					changes,
					refs: message.refs ?? [],
					details: message.details,
					save,
					...ctx.cardDeps,
				}),
			);
		if (create.length)
			out.push(
				...creationCards({
					items: create,
					created: (message.created ??= {}),
					save,
					openCreated: ctx.settings().openCreatedFiles,
					store: ctx.cardDeps.store,
					open: ctx.cardDeps.open,
				}),
			);
	} catch (error) {
		out.push(note('chat-error', error.message));
	}
	return out;
}

function actionsBar(message, ctx) {
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
	if (message.versions?.length > 1) {
		const at = message.selected ?? message.versions.length - 1;
		const pager = document.createElement('span');
		pager.className = 'chat-versions';
		pager.innerHTML = `<button type="button" aria-label="Previous version" ${at === 0 ? 'disabled' : ''}>‹</button><span>${at + 1} / ${message.versions.length}</span><button type="button" aria-label="Next version" ${at === message.versions.length - 1 ? 'disabled' : ''}>›</button>`;
		const [prev, next] = pager.querySelectorAll('button');
		prev.onclick = () => ctx.switchVersion(message, at - 1);
		next.onclick = () => ctx.switchVersion(message, at + 1);
		bar.append(pager);
	}
	add('copy', 'Copy', async () => {
		await navigator.clipboard.writeText(message.content);
		notify('Copied to the clipboard');
	});
	// Only this page's own messages can be resent or regenerated.
	if (message.remote) return bar;
	if (message.role === 'user') add('edit', 'Edit and resend', () => ctx.edit(message));
	else if (ctx.isLast(message)) add('refresh', 'Regenerate', () => ctx.regenerate(message));
	return bar;
}

export function messageNode(message, ctx) {
	const node = document.createElement('article');
	node.className = `chat-msg from-${message.role}`;
	const own = !message.author || message.author.id === ctx.me();
	node.setAttribute(
		'aria-label',
		message.role === 'user' ? (own ? 'You' : message.author.name) : 'Assistant',
	);
	if (message.role === 'user') {
		const bubble = document.createElement('div');
		bubble.className = 'chat-bubble';
		bubble.textContent = message.content;
		if (message.files?.length) {
			const files = document.createElement('div');
			files.className = 'chat-msg-files';
			files.innerHTML = message.files
				.map(
					(f) => `<span>${f.kind === 'note' ? noteBadge() : badge(f.kind)}${escape(f.name)}</span>`,
				)
				.join('');
			bubble.prepend(files);
		}
		if (message.remote)
			node.append(
				Object.assign(document.createElement('span'), {
					className: 'chat-author',
					textContent: message.author.name,
				}),
			);
		node.append(bubble, actionsBar(message, ctx));
		return node;
	}
	node.innerHTML = `<span class="chat-avatar" aria-hidden="true">${icon('spark')}</span>`;
	const body = document.createElement('div');
	body.className = 'chat-body';
	const meta = document.createElement('div');
	meta.className = 'chat-meta';
	meta.textContent = message.remote
		? `${message.model ?? 'Assistant'} · asked by ${message.author?.name || 'someone'}`
		: (message.model ?? 'Assistant');
	const content = document.createElement('div');
	content.className = 'md';
	if (message.error) {
		content.classList.add('chat-error');
		content.textContent = message.error;
	} else if (message.streaming && !message.content)
		content.innerHTML = '<span class="chat-typing"><i></i><i></i><i></i></span>';
	else content.append(renderMarkdown(message.content));
	// A reply someone else is still writing has no cards or actions yet.
	if (message.streaming) {
		node.classList.add('streaming');
		body.append(meta, content);
	} else {
		body.append(meta, content, ...cards(message, ctx));
		if (message.content || message.error) body.append(actionsBar(message, ctx));
	}
	node.append(body);
	return node;
}
