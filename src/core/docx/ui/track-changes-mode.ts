import { expectDefined } from './defined';
import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import { Mapping, ReplaceStep, Transform } from 'prosemirror-transform';
import { Fragment, Slice, type Mark, type Node as ProseMirrorNode } from 'prosemirror-model';
import { createClientId } from '../../collab/identity';
import { createCollaborationIdGenerator } from './collaboration-identity';
import { isHistoryTransaction } from 'prosemirror-history';
import { trackRunFormatting } from './track-run-formatting';
import { trackParagraphFormatting } from './track-paragraph-formatting';

/** Text removed by the latest tracked cut, so pasting it back records a move. */
interface TrackState {
	lastCut?: { text: string; deletionIds: string[] };
}
interface TrackMeta {
	tracked: true;
	cut?: TrackState['lastCut'];
	pasted?: boolean;
}
export const trackChangesPluginKey = new PluginKey<TrackState>('dve-track-changes');
/** Transactions carrying this meta (remote collaboration steps) are never re-tracked. */
export const REMOTE_TRANSACTION_META = 'dve-remote';

function withMark(fragment: Fragment, mark: Mark): Fragment {
	const children: ProseMirrorNode[] = [];
	fragment.forEach((node) => {
		if (node.isText || node.type.name === 'hardBreak')
			children.push(node.mark(mark.addToSet(node.marks)));
		else children.push(node.copy(withMark(node.content, mark)));
	});
	return Fragment.fromArray(children);
}

interface Segment {
	from: number;
	to: number;
	ownInsertion: boolean;
}
function segmentsOf(doc: ProseMirrorNode, from: number, to: number, author: string): Segment[] {
	const segments: Segment[] = [];
	doc.nodesBetween(from, to, (node, pos) => {
		if (!node.isText && node.type.name !== 'hardBreak') return true;
		const start = Math.max(pos, from);
		const end = Math.min(pos + node.nodeSize, to);
		if (end <= start) return true;
		const mark = node.marks.find((item) => item.type.name === 'insertion');
		const ownInsertion = Boolean(mark && mark.attrs.author === author);
		const last = segments.at(-1);
		if (last && last.ownInsertion === ownInsertion && last.to === start) last.to = end;
		else segments.push({ from: start, to: end, ownInsertion });
		return true;
	});
	return segments;
}

/**
 * Replays one `ReplaceStep` so that removed content the current author didn't just insert stays in
 * the document with a `deletion` mark instead of being removed, and newly inserted content gets an
 * `insertion` mark. Deleting one's own still-pending insertion removes it outright (Word's behavior).
 */
interface TrackedChange {
	deletionIds: string[];
	deletedText: string;
	insertionId?: string;
	insertedText: string;
}

function applyTrackedReplace(
	transform: Transform,
	{ from, to, slice }: Pick<ReplaceStep, 'from' | 'to' | 'slice'>,
	author: string,
	date: string,
	nextRevisionId: () => string,
): TrackedChange {
	const schema = transform.doc.type.schema;
	const change: TrackedChange = { deletionIds: [], deletedText: '', insertedText: '' };
	const segments = segmentsOf(transform.doc, from, to, author);
	const keptLength = segments
		.filter((segment) => !segment.ownInsertion)
		.reduce((sum, segment) => sum + (segment.to - segment.from), 0);
	const insertAt = from + keptLength;
	for (const segment of segments)
		if (!segment.ownInsertion)
			change.deletedText += transform.doc.textBetween(segment.from, segment.to, '\n');
	for (const segment of [...segments].reverse()) {
		if (segment.ownInsertion) transform.delete(segment.from, segment.to);
		else {
			const id = nextRevisionId();
			change.deletionIds.push(id);
			transform.addMark(
				segment.from,
				segment.to,
				expectDefined(schema.marks.deletion, 'deletion mark').create({ author, date, id }),
			);
		}
	}
	if (slice.size) {
		const id = nextRevisionId();
		change.insertionId = id;
		change.insertedText = slice.content.textBetween(0, slice.content.size, '\n');
		const mark = expectDefined(schema.marks.insertion, 'insertion mark').create({
			author,
			date,
			id,
		});
		transform.replace(
			insertAt,
			insertAt,
			new Slice(withMark(slice.content, mark), slice.openStart, slice.openEnd),
		);
	}
	return change;
}

/** Names link both move sides; shared client identity avoids per-second counter collisions. */
const nextMoveName = () => `move${createClientId().replaceAll('-', '')}`;

/** Marks every insertion/deletion carrying one of `ids` as a side of the move `name`. */
function markMove(tr: Transaction, ids: ReadonlySet<string>, name: string): void {
	const move = JSON.stringify({ name });
	tr.doc.descendants((node, pos) => {
		if (!node.isText && node.type.name !== 'hardBreak') return true;
		for (const mark of node.marks) {
			if (
				(mark.type.name !== 'insertion' && mark.type.name !== 'deletion') ||
				!ids.has(String(mark.attrs.id))
			)
				continue;
			tr.removeMark(pos, pos + node.nodeSize, mark.type);
			tr.addMark(pos, pos + node.nodeSize, mark.type.create({ ...mark.attrs, move }));
		}
		return true;
	});
}

/**
 * When enabled, rewrites local editing transactions so insertions/deletions become tracked-change
 * marks (see applyTrackedReplace) instead of directly changing the visible text. Handles the common
 * cases of pure text replacement (typing, IME, backspace/delete, cut, paste) and supported
 * run-formatting mark changes and paragraph formatting. Structural table edits and remote
 * collaboration steps pass through untouched.
 */
export function trackChangesPlugin(
	getAuthor: () => string,
	isEnabled: () => boolean,
	idGenerator = createCollaborationIdGenerator(createClientId()),
): Plugin {
	// Browsers may perform a cut, paste or drop natively (ProseMirror then only sees the DOM change,
	// without its `uiEvent` tag), so the input type of the latest `beforeinput` is kept as well.
	let lastInput: { event: string; at: number } | undefined;
	const INPUT_EVENTS: Record<string, string> = {
		deleteByCut: 'cut',
		insertFromPaste: 'paste',
		insertFromDrop: 'drop',
		deleteByDrag: 'drop',
	};
	return new Plugin<TrackState>({
		props: {
			handleDOMEvents: {
				beforeinput(_view, event) {
					const kind = INPUT_EVENTS[(event as InputEvent).inputType];
					lastInput = kind ? { event: kind, at: Date.now() } : undefined;
					return false;
				},
			},
		},
		key: trackChangesPluginKey,
		state: {
			init: () => ({}),
			apply(tr, state) {
				const meta = tr.getMeta(trackChangesPluginKey) as TrackMeta | undefined;
				if (meta?.cut) return { lastCut: meta.cut };
				if (meta?.pasted) return {};
				return state;
			},
		},
		appendTransaction(transactions, oldState, newState) {
			if (!isEnabled()) return null;
			if (
				transactions.some(
					(tr) =>
						isHistoryTransaction(tr) ||
						tr.getMeta(trackChangesPluginKey) ||
						tr.getMeta(REMOTE_TRANSACTION_META),
				)
			)
				return null;
			const relevant = transactions.filter((tr) => tr.docChanged);
			if (!relevant.length) return null;
			const steps = relevant.flatMap((tr) => tr.steps);
			const author = getAuthor() || 'Author';
			const date = new Date(Date.now()).toISOString();
			if (!steps.length || !steps.every((step) => step instanceof ReplaceStep)) {
				if (newState.doc.attrs.trackFormatting === false) return null;
				const formatting =
					trackRunFormatting(steps, oldState, newState, author, date, () =>
						idGenerator('revision'),
					) ??
					trackParagraphFormatting(steps, oldState, newState, author, date, () =>
						idGenerator('revision'),
					);
				return formatting?.setMeta(trackChangesPluginKey, { tracked: true }) ?? null;
			}
			const transform = new Transform(oldState.doc);
			// Each step's positions refer to the document after the earlier untracked steps; map them
			// back to the original document, then forward through the tracked edits (which keep
			// deleted text in place).
			const untracked = new Mapping();
			const changes = steps.map((step) => {
				const replace = step as ReplaceStep;
				const back = untracked.invert();
				const from = transform.mapping.map(back.map(replace.from, 1), 1);
				const to = Math.max(from, transform.mapping.map(back.map(replace.to, -1), -1));
				untracked.appendMap(replace.getMap());
				return applyTrackedReplace(
					transform,
					{ from, to, slice: replace.slice },
					author,
					date,
					() => idGenerator('revision'),
				);
			});
			if (!transform.docChanged) return null;
			// newState already contains the untracked edit: undo it first so the tracked steps,
			// computed against oldState.doc, apply to the document they were derived from.
			const result = newState.tr;
			const applied = relevant.flatMap((tr) =>
				tr.steps.map((step, index) => ({
					step,
					doc: expectDefined(tr.docs[index], 'document before step'),
				})),
			);
			for (const { step, doc } of [...applied].reverse()) result.step(step.invert(doc));
			for (const step of transform.steps) result.step(step);
			const meta: TrackMeta = { tracked: true };
			const events = new Set(relevant.map((tr) => tr.getMeta('uiEvent')));
			if (lastInput && Date.now() - lastInput.at < 1000) events.add(lastInput.event);
			lastInput = undefined;
			const deletedText = changes.map((change) => change.deletedText).join('');
			const insertedText = changes.map((change) => change.insertedText).join('');
			const deletionIds = changes.flatMap((change) => change.deletionIds);
			const insertionIds = changes.flatMap((change) =>
				change.insertionId ? [change.insertionId] : [],
			);
			const lastCut = trackChangesPluginKey.getState(oldState)?.lastCut;
			const trackMoves = newState.doc.attrs.trackMoves !== false;
			if (trackMoves && events.has('drop') && deletedText && deletedText === insertedText)
				// Dragging selected text records a move, as Word does.
				markMove(result, new Set([...deletionIds, ...insertionIds]), nextMoveName());
			else if (trackMoves && events.has('paste') && lastCut && insertedText === lastCut.text) {
				// Pasting the text of the latest cut completes a move.
				markMove(result, new Set([...lastCut.deletionIds, ...insertionIds]), nextMoveName());
				meta.pasted = true;
			} else if (events.has('cut') && deletedText) meta.cut = { text: deletedText, deletionIds };
			result.setMeta(trackChangesPluginKey, meta);
			return result;
		},
	});
}
