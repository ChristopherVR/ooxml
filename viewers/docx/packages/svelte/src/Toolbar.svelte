<script lang="ts">
 import type { WordToolbarProps } from './runtime';
 let { toolbarProps }: { toolbarProps: WordToolbarProps } = $props();
</script>
<div role="toolbar">
 {#each ['Bold', 'Italic', 'Underline', 'Undo', 'Redo'] as label}
  {@const action = label === 'Bold' ? toolbarProps.onBold : label === 'Italic' ? toolbarProps.onItalic : label === 'Underline' ? toolbarProps.onUnderline : label === 'Undo' ? toolbarProps.onUndo : toolbarProps.onRedo}
  {@const pressed = label === 'Bold' ? toolbarProps.bold : label === 'Italic' ? toolbarProps.italic : label === 'Underline' ? toolbarProps.underline : false}
  <button type="button" disabled={label === 'Undo' ? !toolbarProps.canUndo : label === 'Redo' ? !toolbarProps.canRedo : !toolbarProps.canEdit} aria-pressed={pressed === null ? 'mixed' : pressed} onmousedown={(event) => event.preventDefault()} onclick={() => action()}>{toolbarProps.labels[label] ?? label}</button>
 {/each}
</div>
