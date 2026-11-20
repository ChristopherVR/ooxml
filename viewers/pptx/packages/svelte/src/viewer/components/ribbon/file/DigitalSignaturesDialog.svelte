<script lang="ts">
	/**
	 * File > Info > Digital Signatures: the deck's signature state and one row per signature, all
	 * from the shared `signatureStatusView`. A deck reads as signed only when its signatures verify.
	 */
	import DialogFooter from '../../DialogFooter.svelte';
	import Info from '@lucide/svelte/icons/info';
	import X from '@lucide/svelte/icons/x';
	import type { SignatureStatusView } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../../i18n/context';
	import SignatureStatusIcon from '../../SignatureStatusIcon.svelte';

	const { view, onclose }: { view: SignatureStatusView; onclose: () => void } = $props();
	const t = useTranslator();
</script>

<div class="backdrop">
	<button class="scrim" type="button" aria-label={t('pptx.common.close')} onclick={onclose}></button>
	<div class="panel" role="dialog" aria-modal="true" aria-labelledby="pptx-svelte-signatures-title" data-pptx-signature-state={view.state}>
		<header><div><span class="mark" data-tone={view.tone}><SignatureStatusIcon icon={view.icon} size={17} /></span><h2 id="pptx-svelte-signatures-title">{view.heading}</h2></div><button type="button" aria-label={t('pptx.common.close')} onclick={onclose}><X size={16} aria-hidden="true" /></button></header>
		<div class="body">
			{#if view.present}
				<div class="notice" data-tone={view.tone}><b><SignatureStatusIcon icon={view.icon} /></b><p><strong>{view.label}</strong><small>{view.description}</small></p></div>
				<ul class="rows">
					{#each view.rows as row (row.path)}
						<li data-pptx-signature-row-state={row.state} data-tone={row.tone}>
							<b><SignatureStatusIcon icon={row.icon} /></b>
							<div>
								<p><strong title={row.subject}>{row.signer}</strong><span class="state">{row.label}</span></p>
								{#if row.issuer}<small>{t('pptx.digitalSignatures.issuer')}: {row.issuer}</small>{/if}
								{#if row.signedAt}<small>{view.signedAtCaption}: {row.signedAt}</small>{/if}
							</div>
						</li>
					{/each}
				</ul>
				<div class="notice" data-tone="caution"><b><Info size={14} aria-hidden="true" /></b><p>{view.removedOnSave}</p></div>
			{:else}
				<div class="notice"><b><Info size={14} aria-hidden="true" /></b><p>{t('pptx.digitalSignatures.noSignatures')}</p></div>
			{/if}
		</div>
		<footer><DialogFooter actions={[{ id: 'close', label: t('pptx.common.close') }]} onaction={onclose} /></footer>
	</div>
</div>

<style>
	.backdrop{position:fixed;inset:0;z-index:90;display:grid;place-items:center;background:#0009}.scrim{position:absolute;inset:0;width:100%;height:100%;border:0;background:transparent}.panel{position:relative;width:min(420px,calc(100vw - 32px));border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:12px;background:var(--pptx-card, var(--office-surface, #f3f4f6));color:var(--pptx-card-foreground, var(--office-foreground, #1f2937));box-shadow:0 24px 80px #0008}header,footer{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid var(--pptx-border, var(--office-border, #d1d5db))}header div{display:flex;align-items:center;gap:9px}.mark{display:grid;width:24px;height:24px;place-items:center;border-radius:50%;background:var(--pptx-muted, var(--office-surface, #f3f4f6))}h2,p{margin:0}h2{font-size:14px}.body{display:grid;gap:12px;padding:20px}.notice,.rows li{display:flex;align-items:flex-start;gap:11px;padding:12px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:9px;background:var(--pptx-muted, var(--office-surface, #f3f4f6));font-size:12px}.notice b,.rows b{display:grid;width:20px;height:20px;flex:none;place-items:center;border-radius:50%;border:1px solid currentColor}.notice small,.rows small{display:block;margin-top:5px;opacity:.75}.rows{display:grid;gap:8px;margin:0;padding:0;list-style:none}.rows p{display:flex;flex-wrap:wrap;gap:8px;align-items:baseline}.rows .state{font-size:11px}[data-tone='success']{border-color:#15803d66;color:#15803d}[data-tone='caution']{border-color:#b4530966;color:#b45309}[data-tone='danger']{border-color:#b91c1c66;color:#b91c1c}button{border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:6px;padding:6px 10px;background:var(--pptx-muted, var(--office-surface, #f3f4f6));color:inherit}footer{justify-content:flex-end;border-top:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-bottom:0}@media(max-width:600px){.panel{position:fixed;inset:auto 0 0;width:100%;border-radius:16px 16px 0 0}}
</style>
