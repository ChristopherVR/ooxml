<script lang="ts">
	import { buildRecordRibbon, runRecordRibbonAction } from 'ooxml-ui/pptx';
	import type { CanvasSize, RecordTarget, RibbonCommandRequestEvent } from 'ooxml-ui/pptx';
	import { useTranslator } from '../../../../i18n/context';

	import type { EditorState } from '../../../editor/editor-state.svelte';

	const { editor, canvasSize, onfrombeginning, onfromcurrent }: { editor: EditorState; canvasSize: CanvasSize; onfrombeginning: () => void; onfromcurrent: () => void } = $props();
	const t = useTranslator();
	const target = (): RecordTarget | undefined => editor.editable ? {
		slides: editor.slides,
		slideIndex: editor.currentSlideIndex,
		canvasSize,
		insert: editor.insertElement,
		commitSlides: (slides) => editor.commitSlides(slides, t('pptx.ribbon.tab.record')),
	} : undefined;
	const groups = $derived(buildRecordRibbon(t, target()));
	function request(event: Event): void {
		const id = (event as RibbonCommandRequestEvent).detail.id;
		if (id === 'record.record.fromBeginning') {
			onfrombeginning();
		} else if (id === 'record.record.fromCurrent') {
			onfromcurrent();
		} else if (event.currentTarget instanceof HTMLElement) runRecordRibbonAction(event.currentTarget, id, target, t);
	}
</script>

<pptx-ui-ribbon-section {groups} oncommand-request={request}></pptx-ui-ribbon-section>
