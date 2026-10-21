<script lang="ts">
	import DialogFooter from '../../DialogFooter.svelte';
	import BadgeCheck from '@lucide/svelte/icons/badge-check';
	import CircleAlert from '@lucide/svelte/icons/circle-alert';
	import Info from '@lucide/svelte/icons/info';
	import X from '@lucide/svelte/icons/x';

	import { useTranslator } from '../../../../i18n/context';

	const { hasSignatures, signatureCount, onclose }: { hasSignatures: boolean; signatureCount: number; onclose: () => void } = $props();
	const t = useTranslator();
</script>

<div class="backdrop">
	<button class="scrim" type="button" aria-label={t('pptx.common.close')} onclick={onclose}></button>
	<div class="panel" role="dialog" aria-modal="true" aria-labelledby="pptx-svelte-signatures-title">
		<header><div><span class:signed={hasSignatures}>{#if hasSignatures}<BadgeCheck size={17} aria-hidden="true" />{:else}<CircleAlert size={17} aria-hidden="true" />{/if}</span><h2 id="pptx-svelte-signatures-title">{t('pptx.digitalSignatures.title')}</h2></div><button type="button" aria-label={t('pptx.common.close')} onclick={onclose}><X size={16} aria-hidden="true" /></button></header>
		<div class="body">
			{#if hasSignatures}
				<div class="notice signed"><b><BadgeCheck size={14} aria-hidden="true" /></b><p>{t('pptx.digitalSignatures.signed')}<small>{t('pptx.digitalSignatures.signatureCount', { count: signatureCount })}</small></p></div>
				<div class="notice warning"><b><Info size={14} aria-hidden="true" /></b><p>{t('pptx.digitalSignatures.editWarning')}</p></div>
			{:else}
				<div class="notice"><b><Info size={14} aria-hidden="true" /></b><p>{t('pptx.digitalSignatures.noSignatures')}</p></div>
			{/if}
		</div>
		<footer><DialogFooter actions={[{ id: 'close', label: t('pptx.common.close') }]} onaction={onclose} /></footer>
	</div>
</div>

<style>
	.backdrop{position:fixed;inset:0;z-index:90;display:grid;place-items:center;background:#0009}.scrim{position:absolute;inset:0;width:100%;height:100%;border:0;background:transparent}.panel{position:relative;width:min(420px,calc(100vw - 32px));border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:12px;background:var(--pptx-card, var(--office-surface, #f3f4f6));color:var(--pptx-card-foreground, var(--office-foreground, #1f2937));box-shadow:0 24px 80px #0008}header,footer{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid var(--pptx-border, var(--office-border, #d1d5db))}header div{display:flex;align-items:center;gap:9px}header span{display:grid;width:24px;height:24px;place-items:center;border-radius:50%;background:#92400e55;color:#fbbf24;font-weight:700}header span.signed{background:#16653455;color:#4ade80}h2,p{margin:0}h2{font-size:14px}.body{display:grid;gap:12px;padding:20px}.notice{display:flex;align-items:flex-start;gap:11px;padding:12px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:9px;background:var(--pptx-muted, var(--office-surface, #f3f4f6));font-size:12px}.notice.signed{border-color:#15803d66;background:#14532d33;color:#bbf7d0}.notice.warning{border-color:#b4530966;background:#78350f33;color:#fde68a}.notice b{display:grid;width:20px;height:20px;flex:none;place-items:center;border-radius:50%;border:1px solid currentColor}.notice small{display:block;margin-top:5px;opacity:.75}button{border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:6px;padding:6px 10px;background:var(--pptx-muted, var(--office-surface, #f3f4f6));color:inherit}footer{justify-content:flex-end;border-top:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-bottom:0}@media(max-width:600px){.panel{position:fixed;inset:auto 0 0;width:100%;border-radius:16px 16px 0 0}}
</style>
