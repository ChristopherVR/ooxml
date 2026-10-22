<script lang="ts">
 import { prepareWordToolbar, type ControlGroup } from './runtime';
 let { role, groups }: { role: string; groups: readonly ControlGroup[] } = $props();
</script>
<div {role} use:prepareWordToolbar>
 {#each groups as group, index (index)}
  <span role="group">
   {#each group as control (control.id)}
    {#if control.kind === 'text'}
     <span role="status">{control.text}</span>
    {:else if control.kind === 'select'}
     <select aria-label={control.label} disabled={control.disabled} value={control.value} onchange={(event) => control.change(event.currentTarget.value)}>
      {#each control.options as option (option.value)}
       <option value={option.value} selected={option.value === control.value}>{option.label}</option>
      {/each}
     </select>
    {:else}
     <button type="button" aria-label={control.label} disabled={control.disabled} aria-pressed={control.kind === 'toggle' ? control.pressed : undefined} onmousedown={(event) => event.preventDefault()} onclick={() => control.run()}>{control.label}</button>
    {/if}
   {/each}
  </span>
 {/each}
</div>
