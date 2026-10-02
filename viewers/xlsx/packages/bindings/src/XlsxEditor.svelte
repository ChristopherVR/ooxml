<script lang="ts">
  import { deferredHandle, eventOptions, mountEditor, pickEditorProps, type EditorBinding, type EditorEventHandlers, type EditorProps } from './index';
  type Props = EditorProps & {
    onworkbookchange?: EditorEventHandlers['workbook-change'];
    onworkbookerror?: EditorEventHandlers['workbook-error'];
    onselectionchange?: EditorEventHandlers['selection-change'];
    ondirtychange?: EditorEventHandlers['dirty-change'];
    onready?: EditorEventHandlers['ready'];
  };
  let { workbook, bytes, src, fileName, readOnly = false, locale = 'en', theme = 'auto', authorName = 'Author', showToolbar = true, showFormulaBar = true, hiddenActions = [], themeColors, onworkbookchange, onworkbookerror, onselectionchange, ondirtychange, onready }: Props = $props();
  let binding: EditorBinding | undefined;
  const handle = deferredHandle(() => binding);
  function attach(host: HTMLElement, options: Props) {
    const normalized = (value: Props) => ({ ...pickEditorProps(value),
      ...eventOptions({ 'workbook-change': value.onworkbookchange, 'workbook-error': value.onworkbookerror, 'selection-change': value.onselectionchange, 'dirty-change': value.ondirtychange, ready: value.onready }) });
    binding = mountEditor(host, normalized(options));
    return { update(next: Props) { binding?.update(normalized(next)); }, destroy() { binding?.destroy(); binding = undefined; } };
  }
  export function load(input: Uint8Array | ArrayBuffer, name?: string) { return handle.load(input, name); }
  export function newWorkbook() { handle.newWorkbook(); }
  export function save() { return handle.save(); }
  export function saveBytes(format?: 'xlsx' | 'csv') { return handle.saveBytes(format); }
  export function download(name?: string) { return handle.download(name); }
  export function markClean() { handle.markClean(); }
  export function select(ref: string) { handle.select(ref); }
  export function getSelection() { return handle.getSelection(); }
  export function setActiveSheet(index: number) { handle.setActiveSheet(index); }
  export function isDirty() { return handle.dirty; }
  export function getElement() { return binding?.element; }
</script>
<div use:attach={{ workbook, bytes, src, fileName, readOnly, locale, theme, authorName, showToolbar, showFormulaBar, hiddenActions, themeColors, onworkbookchange, onworkbookerror, onselectionchange, ondirtychange, onready }}></div>
