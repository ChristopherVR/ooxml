<script lang="ts">
 import { toolbarButtons, type SpreadsheetToolbarProps } from './runtime';
 let { toolbarProps }: { toolbarProps: SpreadsheetToolbarProps } = $props();
</script>
<div role="toolbar">
 {#each toolbarButtons(toolbarProps) as button (button.label)}
  <button type="button" disabled={button.disabled} aria-pressed={button.pressed === null ? 'mixed' : button.pressed}
   onmousedown={(e) => e.preventDefault()} onclick={() => { void button.run(); }}>{button.label}</button>
 {/each}
</div>
