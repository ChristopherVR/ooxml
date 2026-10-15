/**
 * The cards under an assistant reply. Nothing the assistant proposes is written until the user
 * presses a card's button. Edits show what they replace, can be picked one by one, are checked
 * against the file revisions the request read, saved as one transaction (the old bytes kept as a
 * version) and can be undone; new files are built by the core and opened in a tab.
 */
import { badge, escape, task } from '../ui.js';
import { buildFile, prepareChanges, validateChanges } from '../ai-operations.js';
import { fileName, outline, summarize, validateCreate } from './protocol.js';
import { wordDiff } from './diff.js';

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

function preview(change, detail) {
	if (change.type === 'cells') {
		const table = el('table', 'ai-cell-diff');
		table.innerHTML = '<thead><tr><th>Cell</th><th>Now</th><th>Proposed</th></tr></thead>';
		const body = table.appendChild(el('tbody'));
		for (const [i, cell] of change.cells.entries()) {
			const row = body.appendChild(el('tr'));
			row.append(el('td', '', cell.address));
			row.appendChild(el('td')).append(el('del', '', detail?.cells?.[i]?.before || '(empty)'));
			row.appendChild(el('td')).append(el('ins', '', cell.input));
		}
		return table;
	}
	const diff = el('div', 'ai-diff');
	if (detail) diff.append(wordDiff(detail.before, detail.after));
	else diff.append(el('ins', '', change.text));
	return diff;
}

/** Edits to attached files. `refs` are the files as the request read them. */
export function proposalCard({
	changes,
	refs,
	details = [],
	store,
	sessions,
	refresh,
	save,
	showVersions,
}) {
	validateChanges(changes, refs);
	const panel = el('div', 'ai-proposal');
	const heading = el('div', 'ai-card-heading');
	heading.innerHTML = `<strong>${changes.length} proposed edit${changes.length === 1 ? '' : 's'}</strong>`;
	const description = el(
		'p',
		'',
		'Review each change. Untick any you do not want, then apply the rest.',
	);
	const list = el('ol', 'ai-change-list');
	const boxes = changes.map((change, index) => {
		const ref = refs.find((d) => d.id === change.documentId);
		// Edits arrive grouped by file; each file is named once, above its edits.
		if (change.documentId !== changes[index - 1]?.documentId) {
			const file = list.appendChild(el('li', 'ai-change-file'));
			file.innerHTML = `${badge(ref.kind)}<strong>${escape(ref.name)}</strong>`;
		}
		const li = list.appendChild(el('li', 'ai-change'));
		const label = li.appendChild(el('label'));
		const box = label.appendChild(el('input'));
		box.type = 'checkbox';
		box.checked = true;
		box.setAttribute('aria-label', `Apply edit ${index + 1} to ${ref.name}`);
		label.append(el('small', '', details[index]?.where ?? `Edit ${index + 1}`));
		li.append(preview(change, details[index]));
		box.onchange = () => {
			li.classList.toggle('skipped', !box.checked);
			apply.disabled = !boxes.some((b) => b.checked);
		};
		return box;
	});
	const details_ = el('details');
	details_.append(el('summary', '', 'Show the operations'));
	details_.append(
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
	const apply = button(changes.length > 1 ? 'Apply selected' : 'Apply changes', true);
	if (changes.length > 1) apply.setAttribute('aria-label', 'Apply changes');
	const dismiss = button('Dismiss');
	actions.append(apply, dismiss);
	panel.append(heading, description, list, details_, actions);
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
			const chosen = changes.filter((_, i) => boxes[i].checked);
			const targets = refs.filter((d) => chosen.some((c) => c.documentId === d.id));
			apply.disabled = true;
			const unlock = sessions.lock(targets.map((d) => d.id));
			try {
				const docs = await current(targets);
				const next = await prepareChanges(chosen, docs);
				await current(targets);
				await store.commit(
					next.map((document) => ({ document, expected: document.revision - 1 })),
					'Assistant edits',
				);
				for (const doc of next) await sessions.reload(doc);
				await refresh();
				save?.({ applied: boxes.map((b) => b.checked) });
				for (const box of boxes) box.disabled = true;
				description.textContent = `Applied ${chosen.length} of ${changes.length}. The previous version of each file was kept, and you can undo this set while the files remain unchanged.`;
				actions.replaceChildren();
				const undo = button('Undo these changes');
				actions.append(undo);
				if (showVersions)
					for (const doc of next) {
						const history = button(next.length > 1 ? `History: ${doc.name}` : 'Version history');
						history.title = `Version history of ${doc.name}`;
						history.onclick = () => showVersions(doc.id);
						actions.append(history);
					}
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
							'Undid assistant edits',
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
export function creationCards({ items, created = {}, store, open, save, openCreated = true }) {
	validateCreate(items);
	return items.map((item, index) => {
		const name = fileName(item);
		const card = el('div', 'ai-create-card');
		const head = el('div', 'ai-card-heading');
		head.innerHTML = `${badge(item.type)}<span><strong>${escape(name)}</strong><small>${escape(summarize(item))}</small></span>`;
		const lines = outline(item).filter((line) => line.trim());
		const outlineList = el('ol', 'ai-outline');
		for (const line of lines.slice(0, 8)) outlineList.append(el('li', '', line));
		if (lines.length > 8) outlineList.append(el('li', 'more', `and ${lines.length - 8} more`));
		const actions = el('div', 'ai-card-actions');
		const make = button(created[index] ? 'Open' : openCreated ? 'Create and open' : 'Create', true);
		actions.append(make);
		card.append(head, ...(lines.length ? [outlineList] : []), actions);
		make.onclick = () =>
			task(async () => {
				const existing = created[index] && (await store.get(created[index]).catch(() => undefined));
				if (existing) return open(existing.id);
				make.disabled = true;
				try {
					const doc = await store.create(name, await buildFile(item));
					created[index] = doc.id;
					save?.({ created });
					make.textContent = 'Open';
					if (openCreated) await open(doc.id);
				} finally {
					make.disabled = false;
				}
			});
		return card;
	});
}
