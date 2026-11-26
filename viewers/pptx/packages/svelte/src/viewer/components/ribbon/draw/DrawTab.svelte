<script lang="ts">
 import { toolForDrawPreset } from 'ooxml-ui/pptx';
 import type { RibbonDrawRequestEvent } from 'ooxml-ui/pptx';
 import { useTranslator } from '../../../../i18n/context';
 import type { EditorState } from '../../../editor/editor-state.svelte';

 const { editor }: { editor: EditorState } = $props();
 const t = useTranslator();
 const state = $derived({
  tool: editor.inkOps.tool, color: editor.inkOps.color, width: editor.inkOps.width,
  editable: editor.editable, recentColors: editor.mruColors, translate: t,
 });
 function request(event: RibbonDrawRequestEvent): void {
  if (!editor.editable) {return;}
  const intent = event.detail;
  switch (intent.kind) {
   case 'tool': editor.inkOps.setTool(intent.value); break;
   case 'width': editor.inkOps.setWidth(intent.value); break;
   case 'color':
    editor.inkOps.setColor(intent.value);
    if (intent.committed) {editor.recordRecentColor(intent.value);}
    break;
   case 'preset':
    editor.inkOps.setColor(intent.value.color);
    editor.inkOps.setWidth(intent.value.width);
    editor.inkOps.setTool(toolForDrawPreset(intent.value.kind));
  }
 }
</script>

<pptx-ui-ribbon-draw {state} ondraw-request={request}></pptx-ui-ribbon-draw>
