/**
 * Shared plumbing for tool executors: the execution context, the write-policy
 * router that decides whether a mutation is staged or applied, and small typed
 * helpers for locating slides / elements inside an updater.
 */

import type { PptxData, PptxElement, PptxSlide } from 'ooxml-core/pptx';

import type { PptxAiBridge, PptxAiSlidesUpdater } from '../bridge';
import type { PptxAiWritePolicy } from '../config';
import type { ProposalCommit, ProposalStore } from '../proposals';

/** Context every tool executor receives. */
export interface AiToolContext {
	bridge: PptxAiBridge;
	proposals: ProposalStore;
	writePolicy: PptxAiWritePolicy;
}

/** A tool executor: validated `input` in, JSON-serialisable output out. */
export type AiToolExecutor = (ctx: AiToolContext, input: unknown) => unknown | Promise<unknown>;

/** Result of routing a write through the policy. */
export interface WriteRouteResult {
	/** True when the change was applied straight to the deck. */
	applied?: boolean;
	/** True when the change was staged for review. */
	staged?: boolean;
	/** Proposal id (present when staged). */
	proposalId?: string;
	/** Whether an explicit user approval is required before applying. */
	requiresApproval?: boolean;
	/** Human-readable change summary. */
	summary: string[];
}

/** A computed change to the deck, on its way through the write policy. */
export interface DeckWrite {
	label: string;
	commit: ProposalCommit;
	before: PptxData;
	after: PptxData;
	/** Destructive tools ask the user under every policy. */
	forceApproval?: boolean;
}

/**
 * Route a write through the active policy. This is the one path to the deck.
 *
 * - `'auto'` applies it now, unless the tool forces approval.
 * - `'stage'` holds it for review; "accept all" may apply it.
 * - `'approve'`, and any `forceApproval` tool such as `delete_slides`, hold it
 *   until the user accepts that one change.
 */
export function routeDeckWrite(ctx: AiToolContext, write: DeckWrite): WriteRouteResult {
	const forced = write.forceApproval === true;
	const requiresApproval = forced || ctx.writePolicy === 'approve';
	const { forceApproval: _forceApproval, ...proposed } = write;
	const proposal = ctx.proposals.propose({ ...proposed, requiresApproval });
	if (ctx.writePolicy === 'auto' && !forced) {
		if (!ctx.proposals.apply(proposal.id)) {
			const reason = ctx.proposals.get(proposal.id)?.conflict;
			ctx.proposals.revert(proposal.id);
			throw new Error(reason ?? `${write.label} could not be applied.`);
		}
		return { applied: true, summary: proposal.summary };
	}
	return { staged: true, proposalId: proposal.id, requiresApproval, summary: proposal.summary };
}

/** Route a slides updater through the active write policy. */
export function routeWrite(
	ctx: AiToolContext,
	label: string,
	updater: PptxAiSlidesUpdater,
	forceApproval = false,
): WriteRouteResult {
	const slides = ctx.bridge.getSlides();
	return routeDeckWrite(ctx, {
		label,
		commit: 'slides',
		before: { slides } as PptxData,
		after: { slides: updater(structuredClone(slides)) } as PptxData,
		forceApproval,
	});
}

/** Assert a slide index is in range, returning the slide. */
export function requireSlide(slides: PptxSlide[], slideIndex: number): PptxSlide {
	if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= slides.length) {
		throw new Error(`Slide index ${slideIndex} out of range (0-${slides.length - 1}).`);
	}
	return slides[slideIndex];
}

/** Find an element by id on a slide, or throw. */
export function requireElement(slide: PptxSlide, elementId: string): PptxElement {
	const el = slide.elements.find((e) => e.id === elementId);
	if (!el) {
		throw new Error(`Element '${elementId}' not found on slide ${slide.slideNumber}.`);
	}
	return el;
}

/** Generate a fresh element id. */
export function newElementId(): string {
	return `el-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

/** Generate a fresh slide id. */
export function newSlideId(): string {
	return `slide-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
}
