<script lang="ts">
	/**
	 * ExportProgressModal: the progress modal shown while a multi-slide export (PDF / GIF / WebM)
	 * runs. A thin adapter over the shared `pptx-ui-export-progress`, which owns the dialog and
	 * progress bar semantics, focus and the look, and is deliberately not dismissable: the export
	 * ends by completing, erroring or Cancel. `oncancel` aborts cooperatively (the export loops
	 * check the signal between slides). It stays mounted while closed, so it hands focus back
	 * when an export ends.
	 */
	import type { ExportProgressViewState } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../i18n/context';

	const {
		open,
		title,
		progress,
		statusMessage,
		oncancel,
	}: {
		open: boolean;
		title: string;
		progress: number;
		statusMessage?: string;
		oncancel: () => void;
	} = $props();

	const t = useTranslator();

	const state = $derived.by<ExportProgressViewState>(() => {
		// Read here so a language switch hands the element a new state and its labels follow.
		void t('pptx.export.cancel');
		return { open, title, progress, status: statusMessage, translate: (key) => t(key) };
	});
</script>

<pptx-ui-export-progress {state} onexport-cancel={() => oncancel()}></pptx-ui-export-progress>
