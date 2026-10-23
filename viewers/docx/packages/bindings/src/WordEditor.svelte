<script lang="ts">
  import { deferredHandle, mountEditor, svelteEditorOptions, type EditorBinding, type EditorEventHandlers, type EditorHandle, type EditorOptions, type EditorProps } from './index';
  type Props = EditorProps & {
    ondocumentchange?: EditorEventHandlers['document-change'];
    ondocumenterror?: EditorEventHandlers['document-error'];
    ondocumentwarning?: EditorEventHandlers['document-warning'];
    onpagechange?: EditorEventHandlers['page-change'];
    ondirtychange?: EditorEventHandlers['dirty-change'];
    onfilecommand?: EditorEventHandlers['file-command'];
    onreadonlychange?: EditorEventHandlers['readonly-change'];
    onribbonaction?: EditorEventHandlers['ribbon-action'];
    onribboncustomize?: EditorEventHandlers['ribbon-customize'];
    onribbonaddin?: EditorEventHandlers['office-ribbon-add-in'];
    oncollaborationchange?: EditorEventHandlers['collaboration-change'];
  };
  let { documentModel, fileName, reviewAuthor, themeColors, readOnly = false, locale = 'en', theme = 'auto', showThumbnails = false, showToolbar = true, hiddenActions = [], ribbonAddIns, collaboration, customization, storage, ondocumentchange, ondocumenterror, ondocumentwarning, onpagechange, ondirtychange, onfilecommand, onreadonlychange, onribbonaction, onribboncustomize, onribbonaddin, oncollaborationchange }: Props = $props();
  let binding: EditorBinding | undefined;
  const handle = deferredHandle(() => binding);
  // Live `element` and `dirty`, the same handle vocabulary as the other frameworks (`EDITOR_HANDLE_KEYS`).
  let mountedElement = $state.raw<EditorHandle['element'] | undefined>();
  let dirtyState = $state(false);
  const element = $derived(mountedElement);
  const dirty = $derived(dirtyState);
  function attach(host: HTMLElement, options: Props) {
    // Callback props are the lower-case options (`ondocumentchange`); `dirty` is tracked on the way.
    const tracked = (value: Props): EditorOptions => ({ ...svelteEditorOptions(value), onDirtyChange: (next: boolean) => { dirtyState = next; value.ondirtychange?.(next); } });
    binding = mountEditor(host, tracked(options));
    mountedElement = binding.element;
    dirtyState = binding.dirty;
    return { update(next: Props) { binding?.update(tracked(next)); }, destroy() { binding?.destroy(); binding = undefined; mountedElement = undefined; dirtyState = false; } };
  }
  export function load(input: Uint8Array | ArrayBuffer) { return handle.load(input); }
  export function save() { return handle.save(); }
  export function download(fileName?: string) { return handle.download(fileName); }
  export function markClean() { handle.markClean(); dirtyState = handle.dirty; }
  /** @deprecated Read the dirty property instead; kept so existing callers keep working. */
  export function isDirty() { return handle.dirty; }
  export { element, dirty };
</script>
<div use:attach={{ documentModel, fileName, reviewAuthor, themeColors, readOnly, locale, theme, showThumbnails, showToolbar, hiddenActions, ribbonAddIns, collaboration, customization, storage, ondocumentchange, ondocumenterror, ondocumentwarning, onpagechange, ondirtychange, onfilecommand, onreadonlychange, onribbonaction, onribboncustomize, onribbonaddin, oncollaborationchange }}></div>
