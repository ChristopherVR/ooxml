<script lang="ts">
 import { documentCanvas, type CanvasProps } from './runtime';
 let { mount, className }: CanvasProps = $props();
</script>
<div class={className} use:documentCanvas={{ mount }}></div>
