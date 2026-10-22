<script lang="ts">
	/**
	 * Desktop status bar, docked beneath the slide/notes area. A thin adapter around
	 * the shared `pptx-ui-status-bar`: it maps viewer state onto the element and routes
	 * typed intents to the callbacks. Kept independent from the top toolbar so
	 * read-only viewers retain PowerPoint's navigation and zoom chrome.
	 */
	import { dispatchStatusBarIntent, statusBarState } from 'ooxml-ui/pptx';
	import type { Snippet } from 'svelte';
	import { useTranslator } from '../../i18n/context';
	import type { AutosaveStatus } from '../state/autosave.svelte';

	const {
		current,
		total,
		zoomPercent,
		isDirty,
		autosaveStatus,
		lastSavedAt,
		showNotes = false,
		notesExpanded = false,
		isFullscreen = false,
		slideSorterActive = false,
		onzoomin,
		onzoomout,
		onzoomfit,
		onfullscreen,
		onnotestoggle,
		onnormal,
		onslidesorter,
		collaborationSlot,
		hideZoom = false,
		hideFullscreen = false,
	}: {
		current: number;
		total: number;
		zoomPercent: number;
		isDirty: boolean;
		autosaveStatus?: AutosaveStatus;
		/** Epoch ms of the last successful autosave, for "Saved n minutes ago". */
		lastSavedAt?: number | null;
		showNotes?: boolean;
		notesExpanded?: boolean;
		isFullscreen?: boolean;
		slideSorterActive?: boolean;
		onzoomin: () => void;
		onzoomout: () => void;
		onzoomfit: () => void;
		onfullscreen: () => void;
		onnotestoggle?: () => void;
		onnormal?: () => void;
		onslidesorter?: () => void;
		collaborationSlot?: Snippet;
		/** Hide the zoom cluster (`hiddenActions: ['zoom']`). */
		hideZoom?: boolean;
		/** Hide the Slide Show toggle (`hiddenActions: ['fullscreen']`). */
		hideFullscreen?: boolean;
	} = $props();

	const t = useTranslator();
	const state = $derived(
		statusBarState({
			slideCount: total,
			activeSlideIndex: current,
			dirty: isDirty,
			autosave: autosaveStatus
				? { state: autosaveStatus, timestamp: lastSavedAt ?? undefined }
				: undefined,
			zoomPercent: hideZoom ? undefined : zoomPercent,
			notesExpanded,
			// "Normal" is active whenever neither the slide sorter nor the slideshow is up.
			viewMode: isFullscreen ? 'slideShow' : slideSorterActive ? 'sorter' : 'normal',
			handles: { notes: showNotes, sorter: Boolean(onslidesorter), slideShow: !hideFullscreen },
			translate: t,
		}),
	);
	function request(event: Event): void {
		dispatchStatusBarIntent(event, {
			notes: onnotestoggle,
			normal: onnormal,
			sorter: onslidesorter,
			slideShow: onfullscreen,
			zoomOut: onzoomout,
			zoomFit: onzoomfit,
			zoomIn: onzoomin,
		});
	}
</script>

<!--
	No landmark role here, matching the other four bindings: the row is a mixed
	region of live readouts and controls, not a toolbar. The class lets the viewer
	hide the bar on phones, where the mobile bottom bar replaces it.
-->
<pptx-ui-status-bar class="pptx-svelte-statusbar" {state} onstatus-request={request}>
	{#if collaborationSlot}
		<div slot="collaboration" style="display:contents">{@render collaborationSlot()}</div>
	{/if}
</pptx-ui-status-bar>
