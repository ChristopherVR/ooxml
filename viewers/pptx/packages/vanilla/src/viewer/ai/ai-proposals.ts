/**
 * Render the staged-proposal review cards: one card per pending write with its
 * human-readable diff summary and Accept / Reject buttons, plus an Accept-all
 * action when more than one is staged. Actions call back into the panel so it
 * can apply / revert through the {@link ProposalStore} and re-render (proposal
 * changes do not flow through the chat snapshot).
 */

import type { ProposalView } from 'ooxml-ui/pptx/ai';
import { canAcceptAllProposals, humanizeDiffLine } from 'ooxml-ui/pptx/ai';

import type { Translator } from '../i18n';
import { createEl } from '../render';
import { createIcon } from '../ui/icons';

export interface ProposalCallbacks {
	accept(id: string): void;
	reject(id: string): void;
	acceptAll(): void;
}

/**
 * Replace the container's children with the staged-proposal cards, or, when a
 * load dropped them (`discarded`), with a notice saying so.
 */
export function renderProposals(
	doc: Document,
	container: HTMLElement,
	proposals: ProposalView[],
	t: Translator,
	callbacks: ProposalCallbacks,
	discarded = false,
): void {
	container.replaceChildren();
	container.hidden = proposals.length === 0 && !discarded;
	if (proposals.length === 0) {
		if (discarded) {
			const notice = createEl(doc, 'p', 'pptxv-ai-proposals-discarded');
			notice.setAttribute('role', 'status');
			notice.dataset.testid = 'pptx-ai-proposals-discarded';
			notice.textContent = t('pptx.ai.proposalsDiscarded');
			container.appendChild(notice);
		}
		return;
	}

	const header = createEl(doc, 'div', 'pptxv-ai-proposals-header');
	const title = createEl(doc, 'span', 'pptxv-ai-proposals-title');
	title.textContent = t('pptx.ai.pendingChanges', { count: proposals.length });
	header.appendChild(title);
	if (canAcceptAllProposals(proposals)) {
		const acceptAll = createEl(doc, 'button', 'pptxv-ai-proposal-btn is-accept-all');
		acceptAll.type = 'button';
		acceptAll.textContent = t('pptx.ai.acceptAll');
		acceptAll.addEventListener('click', () => callbacks.acceptAll());
		header.appendChild(acceptAll);
	}
	container.appendChild(header);

	for (const proposal of proposals) {
		container.appendChild(renderProposalCard(doc, proposal, t, callbacks));
	}
}

function renderProposalCard(
	doc: Document,
	proposal: ProposalView,
	t: Translator,
	callbacks: ProposalCallbacks,
): HTMLElement {
	const card = createEl(doc, 'div', 'pptxv-ai-proposal');
	card.dataset.proposalId = proposal.id;

	const eyebrow = createEl(doc, 'div', 'pptxv-ai-proposal-eyebrow');
	eyebrow.textContent = t('pptx.ai.proposedChange');
	card.appendChild(eyebrow);

	const label = createEl(doc, 'div', 'pptxv-ai-proposal-label');
	label.textContent = proposal.label;
	card.appendChild(label);

	// Plain-language, human-readable diff lines. The full description is shown
	// (never truncated); the list scrolls rather than clipping when it is long.
	if (proposal.summary.length > 0) {
		const list = createEl(doc, 'ul', 'pptxv-ai-proposal-summary');
		for (const line of proposal.summary) {
			const item = createEl(doc, 'li');
			item.textContent = humanizeDiffLine(line);
			list.appendChild(item);
		}
		card.appendChild(list);
	}

	if (proposal.requiresApproval && !proposal.conflict) {
		const notice = createEl(doc, 'div', 'pptxv-ai-proposal-notice');
		notice.textContent = t('pptx.ai.proposalNeedsApproval');
		card.appendChild(notice);
	}
	if (proposal.conflict) {
		const notice = createEl(doc, 'div', 'pptxv-ai-proposal-notice is-conflict');
		notice.setAttribute('role', 'alert');
		notice.textContent = t('pptx.ai.proposalConflict');
		card.appendChild(notice);
	}

	const actions = createEl(doc, 'div', 'pptxv-ai-proposal-actions');
	const accept = createEl(doc, 'button', 'pptxv-ai-proposal-btn is-accept');
	accept.type = 'button';
	accept.append(createIcon(doc, 'check'), doc.createTextNode(t('pptx.ai.accept')));
	accept.addEventListener('click', () => callbacks.accept(proposal.id));
	const reject = createEl(doc, 'button', 'pptxv-ai-proposal-btn is-reject');
	reject.type = 'button';
	reject.append(createIcon(doc, 'close'), doc.createTextNode(t('pptx.ai.reject')));
	reject.addEventListener('click', () => callbacks.reject(proposal.id));
	actions.append(accept, reject);
	card.appendChild(actions);

	return card;
}
