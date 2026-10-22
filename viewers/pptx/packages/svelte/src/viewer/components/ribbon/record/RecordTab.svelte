<script lang="ts">
	import { buildRecordRibbon } from 'ooxml-ui/pptx';
	import type { RibbonCommandRequestEvent } from 'ooxml-ui/pptx';
	import { useTranslator } from '../../../../i18n/context';

	const { onfrombeginning, onfromcurrent }: { onfrombeginning: () => void; onfromcurrent: () => void } = $props();
	const t = useTranslator();
	const groups = $derived(buildRecordRibbon(t));
	function request(event: Event): void {
		const id = (event as RibbonCommandRequestEvent).detail.id;
		if (id === 'record.record.fromBeginning') {
			onfrombeginning();
		} else if (id === 'record.record.fromCurrent') {
			onfromcurrent();
		}
	}
</script>

<pptx-ui-ribbon-section {groups} oncommand-request={request}></pptx-ui-ribbon-section>
