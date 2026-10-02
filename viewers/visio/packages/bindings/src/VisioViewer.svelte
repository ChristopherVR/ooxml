<script lang="ts">
  import { mountFrameworkViewer, viewerHandle, viewerOptions, type MountedViewer, type ViewerProps, type VsdxSource } from './common.js';
  type Props = ViewerProps & { class?: string; style?: string };
  let { document, pageIndex, zoom, showToolbar, events, class: className, style }: Props = $props();
  const options = $derived({ document, pageIndex, zoom, showToolbar, events } satisfies Required<ViewerProps>);
  let binding: MountedViewer | undefined;
  const handle = viewerHandle(() => binding);
  function attach(host: HTMLElement, props: ViewerProps) {
    const mounted = mountFrameworkViewer(host, viewerOptions(props));
    binding = mounted;
    return {
      update(next: ViewerProps) { mounted.update(viewerOptions(next)); },
      destroy() { binding = undefined; mounted.destroy(); },
    };
  }
  export function getHandle() { return handle; }
  export function load(source: VsdxSource) { return handle.load(source); }
  export function fit() { handle.fit(); }
</script>
<div class={className} {style} use:attach={options}></div>
