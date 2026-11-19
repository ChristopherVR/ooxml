/**
 * {@link ProposalStore} - the staging area for AI-proposed writes.
 *
 * A proposal records what a tool changed, not the deck it produced. Applying
 * one replays those changes on the deck as it is at that moment, so several
 * proposals and the user's later edits all survive. A proposal whose target has
 * since changed is not applied: it stays listed with a conflict for the user to
 * discard or retry. Each applied proposal is one undoable history entry.
 */

import type { PptxData, PptxSlide } from 'ooxml-core/pptx';

import type { PptxAiBridge, PptxAiSlidesUpdater } from './bridge';
import type { AiChangeAnimator } from './change-animator';
import type { DeckChange } from './deck-delta';
import { applyDeckChanges, describeDeckPath, diffDeck, findDeckConflict } from './deck-delta';
import { diffSlides } from './proposals-diff';

/** Which bridge write commits a proposal. */
export type ProposalCommit = 'slides' | 'theme' | 'deck';

/** A write a tool computed, before it is staged. */
export interface ProposedWrite {
	label: string;
	commit: ProposalCommit;
	before: PptxData;
	after: PptxData;
	/** The user must accept this one by itself; "accept all" leaves it pending. */
	requiresApproval: boolean;
}

/** Public (serialisable) view of a proposal, without its recorded changes. */
export interface ProposalView {
	/** Stable id for accept/revert. */
	id: string;
	/** Short label used as the history-entry name when applied. */
	label: string;
	/** Human-readable change lines. */
	summary: string[];
	/** Epoch ms when staged. */
	createdAt: number;
	/** The user must accept this proposal by itself. */
	requiresApproval: boolean;
	/** Set when applying failed because the target changed after staging. */
	conflict?: string;
}

/** A pending, not-yet-applied write. */
export interface StagedProposal extends ProposalView {
	commit: ProposalCommit;
	changes: DeckChange[];
}

let counter = 0;
function nextId(): string {
	counter += 1;
	return `proposal-${Date.now().toString(36)}-${counter}`;
}

/** The whole deck as the bridge sees it now. Shares objects with the editor: do not mutate. */
export function readLiveDeck(bridge: PptxAiBridge): PptxData {
	const meta = bridge.getDeckMeta();
	const deck = bridge.getDeckData?.();
	return {
		...deck,
		slides: bridge.getSlides(),
		theme: bridge.getTheme() ?? deck?.theme,
		width: meta.width,
		height: meta.height,
	} as PptxData;
}

/** The part of the deck a commit writes; nothing else can conflict or change. */
function scope(deck: PptxData, commit: ProposalCommit): object {
	if (commit === 'slides') {
		return { slides: deck.slides };
	}
	return commit === 'theme' ? { theme: deck.theme } : deck;
}

const MAX_SUMMARY_LINES = 6;

function summarize(write: ProposedWrite, changes: DeckChange[]): string[] {
	if (write.commit === 'slides') {
		return diffSlides(write.before.slides, write.after.slides);
	}
	const lines = changes.flatMap((change) =>
		change.kind === 'set'
			? [`Set ${describeDeckPath(change.path)} to ${JSON.stringify(change.after) ?? 'nothing'}`]
			: [],
	);
	const shown = lines.slice(0, MAX_SUMMARY_LINES).map((line) => line.slice(0, 120));
	return shown.length > 0 ? shown : [write.label];
}

export class ProposalStore {
	private readonly proposals = new Map<string, StagedProposal>();

	constructor(
		private readonly bridge: PptxAiBridge,
		private readonly animator?: AiChangeAnimator,
	) {}

	/** Stage a computed write without touching the deck. */
	propose(write: ProposedWrite): StagedProposal {
		const changes = diffDeck(scope(write.before, write.commit), scope(write.after, write.commit));
		const proposal: StagedProposal = {
			id: nextId(),
			label: write.label,
			summary: summarize(write, changes),
			createdAt: Date.now(),
			requiresApproval: write.requiresApproval,
			commit: write.commit,
			changes,
		};
		this.proposals.set(proposal.id, proposal);
		return proposal;
	}

	/** Stage a slides updater, computing its effect on a copy of the current slides. */
	stage(
		label: string,
		updater: PptxAiSlidesUpdater,
		options: { requiresApproval?: boolean } = {},
	): StagedProposal {
		const slides = this.bridge.getSlides();
		return this.propose({
			label,
			commit: 'slides',
			before: { slides } as PptxData,
			after: { slides: updater(structuredClone(slides)) } as PptxData,
			requiresApproval: options.requiresApproval ?? false,
		});
	}

	private replaySlides(slides: PptxSlide[], changes: DeckChange[]): PptxSlide[] {
		const root = { slides };
		applyDeckChanges(changes, root);
		return root.slides;
	}

	/** Replay a proposal on the current deck. Returns the conflict when it no longer fits. */
	private commit(p: StagedProposal): string | undefined {
		const conflict = findDeckConflict(p.changes, scope(readLiveDeck(this.bridge), p.commit));
		if (conflict) {
			return `${describeDeckPath(conflict.path)} changed after this was proposed.`;
		}
		if (p.commit === 'slides') {
			const before = this.bridge.getSlides();
			const after = this.animator ? this.replaySlides(structuredClone(before), p.changes) : null;
			this.bridge.applySlidesUpdate((slides) => this.replaySlides(slides, p.changes), p.label);
			if (after) {
				this.animator?.publish(before, after);
			}
		} else if (p.commit === 'theme') {
			const root = { theme: structuredClone(this.bridge.getTheme() ?? {}) };
			applyDeckChanges(p.changes, root);
			this.bridge.applyTheme(root.theme);
		} else {
			if (!this.bridge.applyDeckData) {
				return 'This viewer cannot apply presentation-level changes.';
			}
			this.bridge.applyDeckData((deck) => {
				applyDeckChanges(p.changes, deck);
				return deck;
			}, p.label);
		}
		return undefined;
	}

	private settle(p: StagedProposal): boolean {
		const conflict = this.commit(p);
		if (conflict) {
			p.conflict = conflict;
			return false;
		}
		this.proposals.delete(p.id);
		return true;
	}

	private ordered(): StagedProposal[] {
		return [...this.proposals.values()].sort((a, b) => a.createdAt - b.createdAt);
	}

	/** List staged proposals (oldest first), without their recorded changes. */
	list(): ProposalView[] {
		return this.ordered().map(({ commit: _commit, changes: _changes, ...view }) => view);
	}

	/** Number of staged proposals. */
	get size(): number {
		return this.proposals.size;
	}

	/** Look up one proposal's serialisable view. */
	get(id: string): ProposalView | undefined {
		return this.list().find((view) => view.id === id);
	}

	/**
	 * Replay a staged proposal on the live deck as a single history entry and
	 * remove it. Returns `false` when the id is unknown or the proposal conflicts;
	 * a conflicting proposal stays listed with its `conflict` set.
	 */
	apply(id: string): boolean {
		const p = this.proposals.get(id);
		return p ? this.settle(p) : false;
	}

	/** Drop a staged proposal without applying it. */
	revert(id: string): boolean {
		return this.proposals.delete(id);
	}

	/**
	 * Apply staged proposals in staged order, each as its own history entry.
	 * Proposals that need individual approval stay pending, and so do conflicts.
	 * Returns the number applied.
	 */
	acceptAll(): number {
		let applied = 0;
		for (const proposal of this.ordered()) {
			if (!proposal.requiresApproval && this.settle(proposal)) {
				applied += 1;
			}
		}
		return applied;
	}

	/** Drop all staged proposals without applying them. */
	clear(): void {
		this.proposals.clear();
	}
}
