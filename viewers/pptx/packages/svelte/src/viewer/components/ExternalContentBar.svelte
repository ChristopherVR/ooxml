<script lang="ts">
	/**
	 * ExternalContentBar: PowerPoint's security bar under the ribbon when the
	 * deck links pictures or media this load left out ("SECURITY WARNING
	 * References to external pictures and other media have been blocked.
	 * [Enable Content]"). A thin adapter around the shared
	 * `pptx-ui-external-content-bar`: the element owns the markup and strings;
	 * this routes its requests. Whether it shows is `ExternalContentState`.
	 */
	import { dispatchExternalContentBarIntent } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../i18n/context';

	const { onenable, onclose }: { onenable: () => void; onclose: () => void } = $props();

	const t = useTranslator();
	const view = $derived({ translate: t });

	function request(event: Event): void {
		dispatchExternalContentBarIntent(event, { enable: onenable, close: onclose });
	}
</script>

<pptx-ui-external-content-bar state={view} onexternal-content-request={request}
></pptx-ui-external-content-bar>
