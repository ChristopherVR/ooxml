<script lang="ts">
	/**
	 * AiProposalsDiscarded: shown where the "Suggested changes" tray would be when
	 * loading another presentation dropped the staged proposals.
	 */
	import { useTranslator } from '../../../i18n/context';

	const t = useTranslator();
</script>

<p
	class="pptx-svelte-ai-proposals-discarded"
	role="status"
	data-testid="pptx-ai-proposals-discarded"
>
	{t('pptx.ai.proposalsDiscarded')}
</p>

<style>
	.pptx-svelte-ai-proposals-discarded {
		margin: 0;
		padding: 8px 12px;
		border-top: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		background: var(--pptx-background, var(--office-background, #ffffff));
		font-size: 12px;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}
</style>
