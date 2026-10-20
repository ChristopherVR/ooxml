<script lang="ts">
	import { buildHelpRibbon } from 'ooxml-ui/pptx';
	import type { RibbonCommandRequestEvent } from 'ooxml-ui/pptx';
	import { useTranslator } from '../../../../i18n/context';
	import type { HelpTabProps } from './help-tab-props';

	const props: HelpTabProps = $props();
	const t = useTranslator();
	// The host wires Settings only while the Options dialog is available.
	const groups = $derived(buildHelpRibbon(t, { options: props.onsettings !== undefined }));
	function request(event: Event): void {
		switch ((event as RibbonCommandRequestEvent).detail.id) {
			case 'help.help.options':
				props.onsettings?.();
				break;
			case 'help.help.keyboardShortcuts':
				props.onshortcuts();
				break;
			case 'help.help.accessibility':
				props.onaccessibility();
				break;
		}
	}
</script>

<pptx-ui-ribbon-section {groups} oncommand-request={request}></pptx-ui-ribbon-section>
