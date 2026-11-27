<script lang="ts">
 /**
  * TransitionsTab: a thin adapter over the shared `pptx-ui-ribbon-transitions`
  * view. The draft is derived from the ACTIVE SLIDE, every control commits the
  * edited fields through `EditorState.transitionOps.applyChange`, Preview
  * replays the transition on the stage without writing, and sound picks are raw
  * patches through `transitionOps.applyChange`.
  */
 import type { RibbonTransitionsRequestEvent } from 'ooxml-ui/pptx';
 import {
  applyRibbonTransitionDraft,
  BUILT_IN_SOUNDS,
  playDeckSound,
  playSlideTransitionPreview,
  readRibbonTransitionDraft,
  ribbonTransitionsIntentUpdates,
  ribbonTransitionsSoundChange,
  ribbonTransitionStockSoundUrl,
 } from 'ooxml-ui/pptx';

 import { useTranslator } from '../../../../i18n/context';
 import type { EditorState } from '../../../editor/editor-state.svelte';
 import type { ChromeUiState } from '../../../state/chrome-ui.svelte';

 const { editor, chromeUi }: { editor: EditorState; chromeUi?: ChromeUiState } = $props();
 const t = useTranslator();
 const slide = $derived(editor.slides[editor.currentSlideIndex]);
 const draft = $derived(readRibbonTransitionDraft(slide));
 const state = $derived({
  draft,
  transition: slide?.transition,
  editable: editor.editable,
  inspectorOpen: Boolean(chromeUi?.inspectorOpen),
  translate: t,
 });

 function request(event: RibbonTransitionsRequestEvent): void {
  const intent = event.detail;
  const patch = ribbonTransitionsIntentUpdates(intent, draft);
  if (patch) {
   editor.transitionOps.applyChange(patch);
   return;
  }
  switch (intent.kind) {
   case 'preview':
    playSlideTransitionPreview(slide?.transition ?? applyRibbonTransitionDraft(undefined, draft), document);
    break;
   case 'applyToAll':
    editor.transitionOps.applyRibbonDraft(draft, true);
    break;
   case 'inspector':
    chromeUi?.toggleInspector();
    break;
   case 'soundPreview': {
    const url = ribbonTransitionStockSoundUrl(slide?.transition);
    if (url) {playDeckSound(url, BUILT_IN_SOUNDS);}
    break;
   }
   default:
    void ribbonTransitionsSoundChange(intent).then((change) => {
     if (change) {editor.transitionOps.applyChange(change);}
     return undefined;
    });
  }
 }
</script>

<pptx-ui-ribbon-transitions {state} ontransitions-request={request}></pptx-ui-ribbon-transitions>
