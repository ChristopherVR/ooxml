<script lang="ts">
	/**
	 * SignatureStatusBadge: the tab-row badge of a signed deck, drawn from the shared
	 * `signatureStatusView`. It reads "Signed" only when the signatures verify, and renders
	 * nothing for an unsigned deck. Clicking it opens the Digital Signatures dialog.
	 */
	import type { SignatureStatusView } from 'ooxml-ui/pptx';
	import SignatureStatusIcon from './SignatureStatusIcon.svelte';

	const { view, onclick }: { view: SignatureStatusView; onclick?: () => void } = $props();
</script>

{#if view.present}
	<button
		type="button"
		class="pptx-svelte-signature-badge"
		data-tone={view.tone}
		data-pptx-signature-state={view.state}
		title={view.description}
		aria-label={view.description}
		disabled={!onclick}
		{onclick}
	>
		<SignatureStatusIcon icon={view.icon} size={13} />
		<span>{view.label}</span>
	</button>
{/if}

<style>
	.pptx-svelte-signature-badge {
		display: inline-flex;
		align-items: center;
		gap: 4px;
		padding: 2px 8px;
		border: 1px solid currentColor;
		border-radius: 6px;
		background: transparent;
		font: inherit;
		font-size: 11px;
		font-weight: 500;
		white-space: nowrap;
		cursor: pointer;
	}

	.pptx-svelte-signature-badge:disabled {
		cursor: default;
	}

	.pptx-svelte-signature-badge[data-tone='success'] {
		color: #15803d;
		background: #16a34a1f;
	}

	.pptx-svelte-signature-badge[data-tone='caution'] {
		color: #b45309;
		background: #d977061f;
	}

	.pptx-svelte-signature-badge[data-tone='danger'] {
		color: #b91c1c;
		background: #dc26261f;
	}

	.pptx-svelte-signature-badge[data-tone='neutral'] {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
		background: var(--pptx-muted, var(--office-surface, #f3f4f6));
	}
</style>
