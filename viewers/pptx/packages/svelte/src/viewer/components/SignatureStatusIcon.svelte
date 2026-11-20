<script lang="ts">
	/** The icon of a signature state (`SignatureStatusView.icon`) in this binding's icon set. */
	import type { Component } from 'svelte';
	import ShieldCheck from '@lucide/svelte/icons/shield-check';
	import ShieldExclamationPoint from '@lucide/svelte/icons/shield-exclamation-point';
	import ShieldQuestionMark from '@lucide/svelte/icons/shield-question-mark';
	import ShieldX from '@lucide/svelte/icons/shield-x';
	import type { SignatureIcon } from 'ooxml-ui/pptx';

	const { icon, size = 14 }: { icon: SignatureIcon; size?: number } = $props();
	const ICONS: Record<SignatureIcon, Component<{ size?: number; 'aria-hidden'?: 'true' }>> = {
		'shield-check': ShieldCheck,
		'shield-alert': ShieldExclamationPoint,
		'shield-x': ShieldX,
		'shield-question': ShieldQuestionMark,
	};
	const Icon = $derived(ICONS[icon]);
</script>

<Icon {size} aria-hidden="true" />
