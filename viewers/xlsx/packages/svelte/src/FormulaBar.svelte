<script lang="ts">
 import { prepareBlock } from './runtime';
 const themeBlock = (host: HTMLElement) => prepareBlock(host, 'formula-bar');
 import { selectReference, type SpreadsheetToolbarProps } from './runtime';
 let { toolbarProps: toolbar }: { toolbarProps: SpreadsheetToolbarProps } = $props();
 let invalid = $state(false); let invalidName = $state(false);
</script>
<div use:themeBlock>
 <input aria-label={toolbar.labels['Name Box']} value={toolbar.selection} aria-invalid={invalidName}
  oninput={() => { invalidName = false; }} onkeydown={(e) => { if (e.key === 'Enter') { e.preventDefault(); invalidName = !selectReference(toolbar, e.currentTarget.value); } }} />
 <input aria-label={toolbar.labels['Formula Bar']} value={toolbar.formula} readonly={!toolbar.canEdit} aria-invalid={invalid}
  oninput={(e) => { invalid = false; toolbar.onFormulaInput(e.currentTarget.value, e.currentTarget.selectionEnd ?? undefined); }}
  onkeydown={(e) => { if (e.key === 'Enter') { e.preventDefault(); invalid = !toolbar.onFormulaChange(e.currentTarget.value); } else if (e.key === 'Escape') { e.preventDefault(); toolbar.onCancelEdit(); invalid = false; } }} />
</div>
