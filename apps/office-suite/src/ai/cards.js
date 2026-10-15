/**
 * The cards under an assistant reply. Nothing the assistant proposes is written until the user
 * presses a card's button: edits are checked against the file revisions the request read, saved
 * as one transaction and can be undone; new files are built by the core and opened in a tab.
 */
import { badge, escape, task } from '../ui.js';
import { buildFile, prepareChanges, validateChanges } from '../ai-operations.js';
import { fileName, outline, summarize, validateCreate } from './protocol.js';

const el = (tag, className, text) => {
	const node = document.createElement(tag);
	if (className) node.className = className;
	if (text !== undefined) node.textContent = text;
	return node;
};
const button = (label, primary) => {
	const b = el('button', primary ? 'primary' : '', label);
	b.type = 'button';
	return b;
};

/** Edits to attached files. `refs` are the files as the request read them. */
export function proposalCard({ changes, refs, store, sessions, refresh, save }) {
	validateChanges(changes, refs);
	const panel = el('div', 'ai-proposal');
	const heading = el('div', 'ai-card-heading');
	heading.innerHTML = `<strong>${changes.length} proposed edit${changes.length === 1 ? '' : 's'}</strong>`;
	const description = el('p', '', 'Review the exact changes below before applying them.');
	const list = el('ul', 'ai-change-list');
	for (const c of changes) {
		const ref = refs.find((d) => d.id === c.documentId);
		const li = el('li');
		li.innerHTML = `${badge(ref.kind)}<span><strong>${escape(ref.name)}</strong> ${escape(
			c.type === 'cells'
				? c.cells.map((cell) => `${cell.address} → ${cell.input}`).join(', ')
				: `→ ${c.text}`,
		)}</span>`;
		list.append(li);
	}
	const details = el('details');
	details.append(el('summary', '', 'Show the operations'));
	details.append(
		el(
			'pre',
			'',
			JSON.stringify(
				changes.map((c) => ({ file: refs.find((d) => d.id === c.documentId)?.name, ...c })),
				null,
				2,
			),
		),
	);
	const actions = el('div', 'ai-card-actions');
	const apply = button('Apply changes', true);
	const dismiss = button('Dismiss');
	actions.append(apply, dismiss);
	panel.append(heading, description, list, details, actions);
	dismiss.onclick = () => panel.remove();
	async function current(docs) {
		const latest = [];
		for (const doc of docs) {
			sessions.sessions.get(doc.id)?.editor?.flush?.();
			const stored = await store.get(doc.id);
			if (!stored || sessions.isDirty(doc.id) || stored.revision !== doc.revision)
				throw new Error(`${doc.name} changed. Ask again using the latest version.`);
			latest.push(stored);
		}
		return latest;
	}
	apply.onclick = () =>
		task(async () => {
			apply.disabled = true;
			const unlock = sessions.lock(refs.map((d) => d.id));
			try {
				const docs = await current(refs);
				const next = await prepareChanges(changes, docs);
				await current(refs);
				await store.commit(next.map((document) => ({ document, expected: document.revision - 1 })));
				for (const doc of next) await sessions.reload(doc);
				await refresh();
				save?.({ applied: true });
				description.textContent =
					'Changes saved. You can undo this set while the files remain unchanged.';
				actions.replaceChildren();
				const undo = button('Undo these changes');
				actions.append(undo);
				undo.onclick = () =>
					task(async () => {
						await current(next);
						const restored = next.map((d) => ({
							...docs.find((old) => old.id === d.id),
							revision: d.revision + 1,
							modified: Date.now(),
						}));
						await store.commit(
							restored.map((document) => ({ document, expected: document.revision - 1 })),
						);
						for (const doc of restored) await sessions.reload(doc);
						await refresh();
						undo.remove();
						description.textContent = 'Changes undone.';
					});
			} catch (error) {
				apply.disabled = false;
				throw error;
			} finally {
				unlock();
			}
		});
	return panel;
}

/** New files. `created` maps a card's index to the id of the file already made from it. */
export function creationCards({ items, created = {}, store, open, save }) {
	validateCreate(items);
	return items.map((item, index) => {
		const name = fileName(item);
		const card = el('div', 'ai-create-card');
		const head = el('div', 'ai-card-heading');
		head.innerHTML = `${badge(item.type)}<span><strong>${escape(name)}</strong><small>${escape(summarize(item))}</small></span>`;
		const lines = outline(item).filter((line) => line.trim());
		const preview = el('ol', 'ai-outline');
		for (const line of lines.slice(0, 8)) preview.append(el('li', '', line));
		if (lines.length > 8) preview.append(el('li', 'more', `and ${lines.length - 8} more`));
		const actions = el('div', 'ai-card-actions');
		const make = button(created[index] ? 'Open' : 'Create and open', true);
		actions.append(make);
		card.append(head, ...(lines.length ? [preview] : []), actions);
		make.onclick = () =>
			task(async () => {
				const existing = created[index] && (await store.get(created[index]));
				if (existing) return open(existing.id);
				make.disabled = true;
				try {
					const doc = await store.create(name, await buildFile(item));
					created[index] = doc.id;
					save?.({ created });
					make.textContent = 'Open';
					await open(doc.id);
				} finally {
					make.disabled = false;
				}
			});
		return card;
	});
}
