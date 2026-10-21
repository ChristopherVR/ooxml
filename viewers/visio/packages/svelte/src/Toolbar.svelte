<script lang="ts">
 import { prepareDiagramBlock } from '../../bindings/src/building-blocks';
 import { toolbarGroups, type ToolbarGroupId } from '../../bindings/src/toolbar-controls';
 function theme(host: HTMLElement) { prepareDiagramBlock(host, 'toolbar'); }
 import type { DiagramToolbarProps } from '../../bindings/src/building-blocks';
 let { toolbarProps, include }: { toolbarProps: DiagramToolbarProps; include?: readonly ToolbarGroupId[] } = $props();
 const groups = $derived(toolbarGroups(toolbarProps, include));
</script>
<div role="toolbar" use:theme>
 {#each groups as group (group.id)}
  <span data-group={group.id}>
   {#each group.controls as control (control.id)}
    {#if control.kind === 'text'}
     <span data-control={control.id}>{control.text}</span>
    {:else if control.kind === 'button'}
     <button type="button" data-control={control.id} disabled={control.disabled} aria-pressed={control.pressed} onclick={() => control.run()}>{control.label}</button>
    {:else}
     <span data-control="search">
      <input type="search" aria-label={control.label} placeholder={control.label} value={control.query} oninput={(event) => control.change(event.currentTarget.value)} />
      <span aria-live="polite">{control.summary}</span>
      <button type="button" data-control={control.previous.id} disabled={control.previous.disabled} onclick={() => control.previous.run()}>{control.previous.label}</button>
      <button type="button" data-control={control.next.id} disabled={control.next.disabled} onclick={() => control.next.run()}>{control.next.label}</button>
     </span>
    {/if}
   {/each}
  </span>
 {/each}
</div>
