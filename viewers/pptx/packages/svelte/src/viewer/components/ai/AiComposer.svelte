<script lang="ts">
	/**
	 * AiComposer: the message input row (textarea + send / stop button). Enter
	 * sends, Shift+Enter inserts a newline. Purely presentational.
	 */
	import Send from '@lucide/svelte/icons/send';
	import Square from '@lucide/svelte/icons/square';

	import { useTranslator } from '../../../i18n/context';

	const {
		isStreaming,
		onsend,
		onstop,
		prefillText = '',
		prefillNonce = 0,
	}: {
		isStreaming: boolean;
		onsend: (text: string) => void;
		onstop: () => void;
		/** One-shot composer prefill text (applied when `prefillNonce` bumps). */
		prefillText?: string;
		/** Bumps on every ask/fix so the composer focuses + applies `prefillText`. */
		prefillNonce?: number;
	} = $props();

	const t = useTranslator();

	let value = $state('');
	const canSend = $derived(value.trim().length > 0);
	// bind:this writes this (invisible to the linter's prefer-const analysis).
	// eslint-disable-next-line prefer-const
	let inputEl = $state<HTMLTextAreaElement | undefined>(undefined);

	// Apply the one-shot prefill whenever the nonce changes (even for empty text,
	// so an "Ask AI about this" that only pins focus still focuses the composer).
	let lastNonce = 0;
	$effect(() => {
		if (prefillNonce !== lastNonce) {
			lastNonce = prefillNonce;
			if (prefillText) {
				value = prefillText;
			}
			inputEl?.focus();
		}
	});

	function submit(): void {
		const trimmed = value.trim();
		if (trimmed.length === 0 || isStreaming) {
			return;
		}
		onsend(trimmed);
		value = '';
	}

	function onKeydown(event: KeyboardEvent): void {
		if (event.key === 'Enter' && !event.shiftKey) {
			event.preventDefault();
			submit();
		}
	}
</script>

<div class="pptx-svelte-ai-composer">
	<div class="pptx-svelte-ai-composer-box">
		<textarea
			class="pptx-svelte-ai-input"
			bind:this={inputEl}
			bind:value
			onkeydown={onKeydown}
			rows="1"
			placeholder={t('pptx.ai.placeholder')}
			aria-label={t('pptx.ai.placeholder')}
		></textarea>
		{#if isStreaming}
			<button
				type="button"
				class="pptx-svelte-ai-send"
				onclick={onstop}
				title={t('pptx.ai.stop')}
				aria-label={t('pptx.ai.stop')}
			>
				<Square size={16} aria-hidden="true" />
			</button>
		{:else}
			<button
				type="button"
				class="pptx-svelte-ai-send"
				class:is-ready={canSend}
				onclick={submit}
				disabled={!canSend}
				title={t('pptx.ai.send')}
				aria-label={t('pptx.ai.send')}
			>
				<Send size={16} aria-hidden="true" />
			</button>
		{/if}
	</div>
</div>

<style>
	.pptx-svelte-ai-composer {
		border-top: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		padding: 8px;
	}

	.pptx-svelte-ai-composer-box {
		display: flex;
		align-items: flex-end;
		gap: 6px;
		padding: 6px 8px;
		border: 1px solid var(--pptx-input, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: var(--pptx-background, var(--office-background, #ffffff));
	}

	.pptx-svelte-ai-composer-box:focus-within {
		border-color: var(--pptx-ring, var(--office-ring, #2563eb));
	}

	.pptx-svelte-ai-input {
		flex: 1;
		min-height: 24px;
		max-height: 128px;
		resize: none;
		border: none;
		background: transparent;
		color: var(--pptx-card-foreground, var(--office-foreground, #1f2937));
		font: inherit;
		font-size: 13px;
		outline: none;
	}

	.pptx-svelte-ai-input::placeholder {
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-ai-send {
		flex-shrink: 0;
		display: inline-flex;
		align-items: center;
		justify-content: center;
		padding: 6px;
		border: none;
		border-radius: 4px;
		background: transparent;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
		cursor: pointer;
	}

	.pptx-svelte-ai-send:hover:not(:disabled) {
		background: var(--pptx-accent, var(--office-selected, #e5e7eb));
		color: var(--pptx-accent-foreground, var(--office-foreground, #1f2937));
	}

	.pptx-svelte-ai-send.is-ready {
		background: var(--pptx-primary, var(--office-accent, #2563eb));
		color: var(--pptx-primary-foreground, var(--office-accent-foreground, #ffffff));
	}

	.pptx-svelte-ai-send.is-ready:hover {
		background: color-mix(in srgb, var(--pptx-primary, var(--office-accent, #2563eb)) 90%, #000);
	}

	.pptx-svelte-ai-send:disabled {
		cursor: default;
		opacity: 0.5;
	}
</style>
