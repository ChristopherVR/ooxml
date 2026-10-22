/**
 * Reviewing proposed edits inside the document itself. Word shows them natively, as tracked
 * changes by "Assistant" that its Review tab accepts or rejects one by one. Excel, PowerPoint and
 * Visio have no tracked changes, so the editor shows the file as it would be, read only, and a
 * panel walks through each change: it selects the change in the editor, and each one can be kept
 * or undone before the kept ones are applied as one saved, undoable version.
 */
import { badge, escape, task } from '../ui.js';
import { prepareChanges, reviewSpreadsheet } from '../ai-operations.js';
import { effectsView, wordDiff } from './diff.js';

const el = (tag, className, text) =>
	Object.assign(document.createElement(tag), {
		...(className ? { className } : {}),
		...(text !== undefined ? { textContent: text } : {}),
	});

/**
 * Opens `ref` (a file the request read) with `items` ({ change, detail, index }) shown in place.
 * `applied(indexes)` reports the card indexes that were saved into the file.
 */
export async function reviewInDocument({ ref, items, store, sessions, open, refresh, applied }) {
	await open(ref.id);
	const session = await sessions.sessions.get(ref.id)?.ready;
	if (!session) throw new Error(`${ref.name} could not be opened.`);
	if (sessions.isDirty(ref.id)) throw new Error(`Save ${ref.name} before reviewing suggestions.`);
	const original = await store.get(ref.id);
	if (original.revision !== ref.revision)
		throw new Error(`${ref.name} changed. Ask again using the latest version.`);
	session.node.querySelector('.review-panel')?.remove();
	const word = ref.kind === 'docx';
	const kept = items.map(() => true);
	// A spreadsheet shows what else the kept changes recalculate, and applies that reviewed proposal.
	const effects = el('div');
	const proposals = {};

	/** The editor shows the document with the kept changes (Word: all of them, as redlines). */
	async function preview() {
		let bytes;
		if (word) {
			const { suggestDocxRunText } = await import('ooxml-core/automation');
			bytes = (
				await suggestDocxRunText(
					original.bytes,
					items.map(({ change }) => change),
				)
			).bytes;
		} else {
			const chosen = items.filter((_, i) => kept[i]).map(({ change }) => change);
			if (ref.kind === 'xlsx') {
				const review = await reviewSpreadsheet(chosen, original);
				delete proposals[ref.id];
				if (review.proposal) proposals[ref.id] = review.proposal;
				effects.replaceChildren(effectsView(review));
			}
			bytes = chosen.length
				? (await prepareChanges(chosen, [original], proposals))[0].bytes
				: original.bytes;
		}
		await sessions.reload({ ...original, bytes });
		session.editor.readOnly?.(!word);
		// The preview is not saved: the session counts as changed until it is applied or discarded.
		if (word) {
			session.version++;
			sessions.update();
		}
	}

	const panel = el('aside', 'review-panel');
	panel.setAttribute('aria-label', 'Assistant suggestions');
	const head = el('div', 'review-head');
	head.innerHTML = `${badge(ref.kind)}<strong>${items.length} suggestion${items.length === 1 ? '' : 's'}</strong>`;
	const collapse = el('button', 'review-collapse', '−');
	collapse.type = 'button';
	collapse.setAttribute('aria-label', 'Collapse suggestions');
	collapse.onclick = () => {
		const folded = panel.classList.toggle('folded');
		collapse.textContent = folded ? '+' : '−';
		collapse.setAttribute('aria-label', folded ? 'Expand suggestions' : 'Collapse suggestions');
	};
	head.append(collapse);
	const help = el(
		'p',
		'review-help',
		word
			? 'Shown in the document as tracked changes by Assistant. Accept or reject each one from the Review tab, then save, or decide for all of them here.'
			: 'The document shows these changes in place. Select one to find it; undo any you do not want, then apply the rest.',
	);
	const list = el('ol', 'review-list');
	items.forEach(({ change, detail }, i) => {
		const li = list.appendChild(el('li', 'review-item'));
		const where = el('button', 'review-where', detail?.where ?? `Change ${i + 1}`);
		where.type = 'button';
		where.disabled = word;
		where.title = word ? 'Find it in the document' : 'Select it in the editor';
		where.onclick = () => task(() => session.editor.reveal?.(change));
		const diff = el('div', 'ai-diff');
		if (change.type === 'cells')
			diff.innerHTML = change.cells
				.map(
					(c, j) =>
						`<span><b>${escape(c.address)}</b> <del>${escape(detail?.cells?.[j]?.before || '(empty)')}</del> <ins>${escape(c.input)}</ins></span>`,
				)
				.join('');
		else if (detail) diff.append(wordDiff(detail.before, detail.after));
		else diff.append(el('ins', '', change.text));
		li.append(where, diff);
		if (!word) {
			const toggle = el('button', 'review-toggle', 'Undo');
			toggle.type = 'button';
			toggle.onclick = () =>
				task(async () => {
					kept[i] = !kept[i];
					li.classList.toggle('skipped', !kept[i]);
					toggle.textContent = kept[i] ? 'Undo' : 'Keep';
					apply.disabled = !kept.some(Boolean);
					await preview();
				});
			li.append(toggle);
		}
	});
	const actions = el('div', 'review-actions');
	const apply = el('button', 'primary', word ? 'Accept all' : 'Apply kept changes');
	const discard = el('button', '', word ? 'Reject all' : 'Discard');
	apply.type = discard.type = 'button';
	actions.append(apply, discard);
	if (word) {
		const leave = el('button', '', 'Close');
		leave.type = 'button';
		leave.title = 'Keep the tracked changes in the document to review in Word';
		leave.onclick = () => finish();
		actions.append(leave);
	}
	panel.append(head, help, list, effects, actions);

	function finish() {
		session.editor.readOnly?.(false);
		panel.remove();
	}
	apply.onclick = () =>
		task(async () => {
			apply.disabled = discard.disabled = true;
			try {
				const latest = await store.get(ref.id);
				if (latest.revision !== original.revision)
					throw new Error(`${ref.name} changed. Ask again using the latest version.`);
				const chosen = items.filter((_, i) => kept[i]);
				const [next] = await prepareChanges(
					chosen.map(({ change }) => change),
					[original],
					proposals,
				);
				await store.commit([{ document: next, expected: original.revision }], 'Assistant edits');
				await sessions.reload(next);
				sessions.update();
				await refresh();
				applied(
					chosen.map(({ index }) => index),
					next,
				);
				finish();
			} finally {
				apply.disabled = discard.disabled = false;
			}
		});
	discard.onclick = () =>
		task(async () => {
			await sessions.reload(await store.get(ref.id));
			sessions.update();
			finish();
		});

	await preview();
	session.node.append(panel);
	if (!word) session.editor.reveal?.(items[0].change);
	return panel;
}
