<script lang="ts">
	import { VIEWER_SHORTCUT_REFERENCE } from 'ooxml-ui/pptx';
	import { useTranslator } from '../../i18n/context';

	const { onclose }: { onclose: () => void } = $props();
	const t = useTranslator();
</script>
<!-- svelte-ignore a11y_no_noninteractive_element_to_interactive_role -->
<aside class="panel" data-pptx-shortcuts-panel="true" role="dialog" tabindex="-1" aria-labelledby="shortcuts-title">
	<header><h2 id="shortcuts-title">{t('pptx.shortcuts.title')}</h2><button type="button" onclick={onclose}>{t('pptx.shortcuts.close')}</button></header>
	<div>{#each VIEWER_SHORTCUT_REFERENCE as shortcut, index}<p class:stripe={index % 2 === 0}><span>{t(shortcut.actionKey)}</span><kbd>{shortcut.shortcut}</kbd></p>{/each}</div>
</aside>
<style>
	.panel{position:absolute;z-index:80;top:56px;right:12px;width:min(384px,calc(100% - 24px));overflow:hidden;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:8px;background:var(--pptx-card, var(--office-surface, #f3f4f6));box-shadow:0 18px 50px #0008}header{display:flex;align-items:center;justify-content:space-between;padding:9px 12px;border-bottom:1px solid var(--pptx-border, var(--office-border, #d1d5db))}h2{margin:0;font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase}button{border:0;border-radius:5px;padding:5px 8px;background:transparent;color:inherit;font-size:11px}button:hover{background:var(--pptx-muted, var(--office-surface, #f3f4f6))}div{max-height:280px;overflow:auto;padding:8px}p{display:flex;align-items:center;justify-content:space-between;gap:14px;margin:0;padding:7px 9px;border-radius:5px;font-size:12px}.stripe{background:var(--pptx-muted, var(--office-surface, #f3f4f6))}kbd{color:var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));font:11px ui-monospace,monospace;white-space:nowrap}
</style>
