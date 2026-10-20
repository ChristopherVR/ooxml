<script lang="ts">
 import type { SpreadsheetToolbarProps } from './runtime';
 let { toolbarProps }: { toolbarProps: SpreadsheetToolbarProps } = $props();
</script>
<div role="tablist" aria-label={toolbarProps.labels['Sheet']}>
 {#each toolbarProps.sheets.filter((s) => s.state === 'visible') as sheet (sheet.index)}
  <button type="button" role="tab" aria-selected={sheet.index === toolbarProps.activeSheet} onclick={() => toolbarProps.onSheetChange(sheet.index)}>{sheet.name}</button>
 {/each}
</div>
