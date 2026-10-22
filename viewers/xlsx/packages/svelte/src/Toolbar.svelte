<script lang="ts">
 import { prepareBlock } from './runtime';
 const themeBlock = (host: HTMLElement) => prepareBlock(host, 'toolbar');
 import { toolbarModel, type SpreadsheetToolbarProps } from './runtime';
 let { toolbarProps }: { toolbarProps: SpreadsheetToolbarProps } = $props();
</script>
<div use:themeBlock role="toolbar">
 {#each toolbarModel(toolbarProps) as item (item.id)}
  {#if item.kind === 'select'}
   <select aria-label={item.label} disabled={item.disabled} value={item.value}
    onchange={(e) => { void item.change(e.currentTarget.value); }}>
    {#each item.options as option (option.value)}
     <option value={option.value} selected={option.value === item.value}>{option.label}</option>
    {/each}
   </select>
  {:else}
   <button type="button" disabled={item.disabled} aria-pressed={item.pressed === null ? 'mixed' : item.pressed}
    onmousedown={(e) => e.preventDefault()} onclick={() => { void item.run(); }}>{item.label}</button>
  {/if}
 {/each}
</div>
