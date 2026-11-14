<script lang="ts">
	/**
	 * FormatSections: the Properties tab as a selection-aware Format pane. The
	 * shared section list for the selection (`resolveFormatSections`), each a
	 * `<pptx-ui-format-section>` whose open state lives in the viewer's Format
	 * pane controller (per category, in memory), with launcher requests
	 * expanded, revealed and focused once rendered. The section bodies are the
	 * binding's existing panels, supplied by the `section` snippet.
	 */
	import type { Snippet } from 'svelte';
	import type { PptxElement } from 'pptx-viewer-core';
	import {
		formatSectionOpen,
		isFormatSectionMixed,
		resolveFormatCategory,
		resolveFormatSections,
		revealFormatSection,
	} from 'ooxml-ui/pptx';
	import type { FormatPaneController, FormatSectionView } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import { controllerState } from '../../state/controller-state';

	const {
		elements,
		pane,
		section,
	}: {
		/** The selection, primary first (empty for the slide). */
		elements: readonly PptxElement[];
		pane: FormatPaneController;
		section: Snippet<[FormatSectionView]>;
	} = $props();
	const t = useTranslator();
	const paneState = $derived(controllerState(pane));
	const views = $derived(resolveFormatSections(elements));
	const category = $derived(resolveFormatCategory(elements));
	let container: HTMLDivElement | undefined = $state();

	interface SectionProps {
		view: FormatSectionView;
		open: boolean;
		mixed: boolean;
		label: string;
		mixedLabel: string;
	}

	type SectionHost = HTMLElement & {
		label: string;
		open: boolean;
		mixed: boolean;
		mixedLabel: string;
	};

	/** Set the element's properties (not attributes) and route its toggle event. */
	function sectionProps(node: HTMLElement, props: SectionProps) {
		let current = props;
		const apply = (next: SectionProps): void => {
			current = next;
			const host = node as SectionHost;
			host.label = next.label;
			host.open = next.open;
			host.mixed = next.mixed;
			host.mixedLabel = next.mixedLabel;
		};
		const listener = (event: Event): void =>
			pane.setOpen(category, current.view.id, (event as CustomEvent<{ open: boolean }>).detail.open);
		apply(props);
		node.addEventListener('office-section-toggle', listener);
		return {
			update: apply,
			destroy: () => node.removeEventListener('office-section-toggle', listener),
		};
	}

	$effect(() => {
		// Track the selection kind, the shown sections and any pending request.
		const currentCategory = category;
		const currentViews = views;
		void paneState.current.request;
		const changed = pane.noteCategory(currentCategory);
		const reveal = pane.takeRequest(currentCategory, currentViews);
		const root = container;
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
	});
</script>

<div
	bind:this={container}
	class="pptx-svelte-format-sections"
	data-pptx-format-sections={category}
	role="group"
	aria-label={t('pptx.formatPane.sections')}
>
	{#each views as view (view.id)}
		<pptx-ui-format-section
			data-pptx-format-section={view.id}
			use:sectionProps={{
				view,
				open: formatSectionOpen(paneState.current, category, view),
				mixed: isFormatSectionMixed(elements, view.id),
				label: t(view.labelKey),
				mixedLabel: t('pptx.formatPane.mixed'),
			}}
		>
			{@render section(view)}
		</pptx-ui-format-section>
	{/each}
</div>

<style>
	.pptx-svelte-format-sections {
		display: flex;
		flex-direction: column;
	}
</style>
