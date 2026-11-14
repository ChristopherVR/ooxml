import {
	FORMAT_SECTION_ATTRIBUTE,
	FORMAT_SECTION_IDS,
	formatSectionOpen,
	isFormatSectionMixed,
	resolveFormatCategory,
	resolveFormatSections,
	revealFormatSection,
} from 'ooxml-ui/pptx';
import type {
	FormatPaneController,
	FormatSectionId,
	FormatSectionView,
	FormatSelectionCategory,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';

type SectionHost = HTMLElement & {
	label: string;
	open: boolean;
	mixed: boolean;
	mixedLabel: string;
};

export interface FormatSections {
	el: HTMLElement;
	/** The `<pptx-ui-format-section>` of `id`. */
	host(id: FormatSectionId): HTMLElement;
	/** A section factory whose cards land in the Format pane section `id`. */
	into(id: FormatSectionId): (label: string) => HTMLElement;
	/** Show, order and label the sections for the selection (after the panels updated). */
	layout(elements: readonly PptxElement[]): void;
	destroy(): void;
}

/**
 * The Properties tab as a Format pane: one `<pptx-ui-format-section>` per
 * shared section id, created once; every existing panel is built into the
 * section `FORMAT_PANEL_SECTIONS` assigns it. `layout` shows the sections the
 * selection has (in the shared order), and the open state and launch requests
 * come from the viewer's Format pane controller.
 */
export function createFormatSections(
	doc: Document,
	t: Translator,
	pane: FormatPaneController,
	scroller: HTMLElement,
	onRequest: () => void,
): FormatSections {
	const el = createEl(doc, 'div', 'pptxv-format-sections');
	el.setAttribute('role', 'group');
	el.setAttribute('aria-label', t('pptx.formatPane.sections'));
	el.style.display = 'flex';
	el.style.flexDirection = 'column';

	let category: FormatSelectionCategory = 'slide';
	let views: FormatSectionView[] = [];
	let shown: FormatSectionView[] = [];
	let laidOut = false;

	const hosts = new Map<FormatSectionId, SectionHost>();
	for (const id of FORMAT_SECTION_IDS) {
		const host = doc.createElement('pptx-ui-format-section') as SectionHost;
		host.setAttribute(FORMAT_SECTION_ATTRIBUTE, id);
		host.hidden = true;
		host.addEventListener('office-section-toggle', (event) => {
			pane.setOpen(category, id, (event as CustomEvent<{ open: boolean }>).detail.open);
		});
		hosts.set(id, host);
		el.appendChild(host);
	}
	const hostOf = (id: FormatSectionId): SectionHost => hosts.get(id)!;

	const sync = (): void => {
		const state = pane.getSnapshot();
		for (const view of views) {
			hostOf(view.id).open = formatSectionOpen(state, category, view);
		}
		if (!state.request || !laidOut) {
			return;
		}
		const reveal = pane.takeRequest(category, shown);
		if (reveal) {
			requestAnimationFrame(() => revealFormatSection(el, reveal.section, reveal));
		}
	};
	// The controller outlives a chrome rebuild (it is per viewer), so an
	// inspector that has been laid out and then removed lets go of it.
	const unsubscribe = pane.subscribe(() => {
		if (laidOut && !el.isConnected) {
			unsubscribe();
			return;
		}
		if (pane.getSnapshot().request) {
			onRequest();
		}
		sync();
	});

	return {
		el,
		host: hostOf,
		into: (id) => (label) => {
			const wrap = createEl(doc, 'div', 'pptxv-inspector-section');
			const caption = createEl(doc, 'h4', 'pptxv-inspector-section-title');
			caption.textContent = label;
			wrap.appendChild(caption);
			hostOf(id).appendChild(wrap);
			return wrap;
		},
		layout(elements) {
			category = resolveFormatCategory(elements);
			views = resolveFormatSections(elements);
			const ids = new Set(views.map((view) => view.id));
			for (const [id, host] of hosts) {
				if (!ids.has(id)) {
					host.hidden = true;
				}
			}
			shown = [];
			views.forEach((view, index) => {
				const host = hostOf(view.id);
				// Move only when out of order, so a focused control is not detached needlessly.
				if (el.children[index] !== host) {
					el.insertBefore(host, el.children[index] ?? null);
				}
				host.label = t(view.labelKey);
				host.mixed = isFormatSectionMixed(elements, view.id);
				host.mixedLabel = t('pptx.formatPane.mixed');
				// A section none of whose panels apply to this element stays out of the way.
				host.hidden = ![...host.children].some((child) => !(child as HTMLElement).hidden);
				if (!host.hidden) {
					shown.push(view);
				}
			});
			laidOut = true;
			const changed = pane.noteCategory(category);
			if (changed && !pane.getSnapshot().request) {
				scroller.scrollTop = 0;
			}
			sync();
		},
		destroy: unsubscribe,
	};
}
