import {
	FORMAT_SECTION_ATTRIBUTE,
	formatSectionOpen,
	isFormatSectionMixed,
	resolveFormatCategory,
	resolveFormatSections,
	revealFormatSection,
} from 'ooxml-ui/pptx';
import type { FormatSectionView, FormatSelectionCategory } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import React, { useEffect, useLayoutEffect, useMemo, useRef } from 'react';

import { useController } from '../../hooks/useController';
import { useTranslation } from '../../i18n/use-translation';
import { useFormatPane } from './format-pane-context';

type SectionHost = HTMLElement & {
	label: string;
	open: boolean;
	mixed: boolean;
	mixedLabel: string;
};

interface FormatSectionProps {
	view: FormatSectionView;
	open: boolean;
	mixed: boolean;
	onToggle: (open: boolean) => void;
	children: React.ReactNode;
}

/** One `<pptx-ui-format-section>`, its properties set natively (React 18 and 19 alike). */
function FormatSection({
	view,
	open,
	mixed,
	onToggle,
	children,
}: FormatSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const ref = useRef<SectionHost>(null);
	const label = t(view.labelKey);
	const mixedLabel = t('pptx.formatPane.mixed');
	useLayoutEffect(() => {
		const host = ref.current;
		if (host) {
			host.label = label;
			host.open = open;
			host.mixed = mixed;
			host.mixedLabel = mixedLabel;
		}
	}, [label, open, mixed, mixedLabel]);
	const toggle = useRef(onToggle);
	toggle.current = onToggle;
	useEffect(() => {
		const host = ref.current;
		if (!host) {
			return;
		}
		const listener = (event: Event) =>
			toggle.current((event as CustomEvent<{ open: boolean }>).detail.open);
		host.addEventListener('office-section-toggle', listener);
		return () => host.removeEventListener('office-section-toggle', listener);
	}, []);
	return (
		<pptx-ui-format-section ref={ref} {...{ [FORMAT_SECTION_ATTRIBUTE]: view.id }}>
			{children}
		</pptx-ui-format-section>
	);
}

interface FormatSectionsProps {
	/** The selection (empty for the slide). */
	elements: readonly PptxElement[];
	/** The body of one section: the binding's existing panels for it. */
	renderSection: (view: FormatSectionView, category: FormatSelectionCategory) => React.ReactNode;
}

/**
 * The Properties tab as a Format pane: the shared section list for the
 * selection, each section's open state in the viewer's Format pane
 * controller, and launch requests revealed once rendered.
 */
export function FormatSections({
	elements,
	renderSection,
}: FormatSectionsProps): React.ReactElement {
	const { t } = useTranslation();
	const pane = useFormatPane();
	const state = useController(pane);
	const views = useMemo(() => resolveFormatSections(elements), [elements]);
	const category = resolveFormatCategory(elements);
	const ref = useRef<HTMLDivElement>(null);

	useLayoutEffect(() => {
		const changed = pane.noteCategory(category);
		const reveal = pane.takeRequest(category, views);
		const root = ref.current;
		if (!root) {
			return;
		}
		if (reveal) {
			// One frame for the section to open and lay out before it is scrolled to.
			requestAnimationFrame(() => revealFormatSection(root, reveal.section, reveal));
		} else if (changed) {
			const scroller = root.closest<HTMLElement>('[data-pptx-chrome="inspector-body"]');
			if (scroller) {
				scroller.scrollTop = 0;
			}
		}
	}, [pane, category, views, state.request]);

	return (
		<div
			ref={ref}
			data-pptx-format-sections={category}
			aria-label={t('pptx.formatPane.sections')}
			role='group'
			className='flex flex-col'
		>
			{views.map((view) => (
				<FormatSection
					key={view.id}
					view={view}
					open={formatSectionOpen(state, category, view)}
					mixed={isFormatSectionMixed(elements, view.id)}
					onToggle={(open) => pane.setOpen(category, view.id, open)}
				>
					{renderSection(view, category)}
				</FormatSection>
			))}
		</div>
	);
}
